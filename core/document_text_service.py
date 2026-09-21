"""Single-pass local text extraction with content-addressed caching."""

from __future__ import annotations

import hashlib
import json
import multiprocessing
import queue
import re
import sqlite3
import time
import zipfile
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Callable, Mapping


TEXT_EXTRACTOR_VERSION = "document-text-v1"


@dataclass(frozen=True)
class TextExtraction:
    text: str
    method: str
    pages: tuple[int, ...]
    diagnostics: Mapping[str, object]
    duration_ms: int
    from_cache: bool


def file_sha256(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_cache(
    connection: sqlite3.Connection | None,
    sha256: str,
    extractor_version: str,
) -> TextExtraction | None:
    if connection is None:
        return None
    try:
        row = connection.execute(
            """
            SELECT text_content, method, pages_json, diagnostics_json, duration_ms
              FROM document_text_cache
             WHERE sha256 = ? AND extractor_version = ?
            """,
            (sha256, extractor_version),
        ).fetchone()
    except sqlite3.OperationalError:
        return None
    if row is None:
        return None
    return TextExtraction(
        text=row[0],
        method=row[1],
        pages=tuple(int(item) for item in json.loads(row[2])),
        diagnostics=json.loads(row[3]),
        duration_ms=int(row[4]),
        from_cache=True,
    )


def _store_cache(
    connection: sqlite3.Connection | None,
    sha256: str,
    extractor_version: str,
    result: TextExtraction,
) -> None:
    if connection is None:
        return
    try:
        connection.execute(
            """
            INSERT INTO document_text_cache (
                sha256, extractor_version, text_content, method, pages_json,
                diagnostics_json, duration_ms
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(sha256, extractor_version) DO UPDATE SET
                text_content = excluded.text_content,
                method = excluded.method,
                pages_json = excluded.pages_json,
                diagnostics_json = excluded.diagnostics_json,
                duration_ms = excluded.duration_ms,
                created_at = datetime('now')
            """,
            (
                sha256,
                extractor_version,
                result.text,
                result.method,
                json.dumps(result.pages),
                json.dumps(dict(result.diagnostics), ensure_ascii=False),
                int(result.duration_ms),
            ),
        )
        connection.commit()
    except sqlite3.OperationalError:
        # Older or read-only databases remain usable; they simply miss the cache.
        return


def _text_is_sufficient(text: str) -> bool:
    compact = re.sub(r"\s+", " ", text or "").strip()
    if len(compact) < 60:
        return False
    return bool(
        re.search(
            r"(?i)\b(?:factura|lectura|contador|propietario|presupuesto|albar[aá]n|"
            r"justificante|informe|periodo|importe|total|iva)\b",
            compact,
        )
    )


def _extract_zip_text(path: Path, max_pages: int) -> TextExtraction:
    started = time.monotonic()
    parts: list[str] = []
    pages: list[int] = []
    with zipfile.ZipFile(path) as archive:
        names = sorted(name for name in archive.namelist() if name.lower().endswith(".txt"))
        for index, name in enumerate(names[:max_pages], start=1):
            parts.append(archive.read(name).decode("utf-8", errors="replace"))
            pages.append(index)
    return TextExtraction(
        "\n".join(parts),
        "embedded_text",
        tuple(pages),
        {},
        int((time.monotonic() - started) * 1000),
        False,
    )


def _default_extract(path: Path, max_pages: int) -> TextExtraction:
    started = time.monotonic()
    if zipfile.is_zipfile(path):
        return _extract_zip_text(path, max_pages)

    text_parts: list[str] = []
    pages: list[int] = []
    try:
        import pdfplumber

        with pdfplumber.open(path) as pdf:
            for number, page in enumerate(pdf.pages[:max_pages], start=1):
                page_text = page.extract_text() or ""
                if page_text.strip():
                    text_parts.append(page_text)
                    pages.append(number)
    except Exception as error:
        text_error = str(error)
    else:
        text_error = ""

    text = "\n".join(text_parts)
    if _text_is_sufficient(text):
        return TextExtraction(
            text,
            "pdf_text",
            tuple(pages),
            {},
            int((time.monotonic() - started) * 1000),
            False,
        )

    # Import lazily to keep the service independently testable and avoid a
    # circular dependency: lector_pdf does not import this module.
    import lector_pdf

    ocr = lector_pdf.extraer_texto_ocr_con_diagnostico(str(path))
    combined = ocr.text.strip() or text
    diagnostics: dict[str, object] = {
        "ocr_status": ocr.status,
        "ocr_detail": ocr.detail,
    }
    if text_error:
        diagnostics["pdf_text_error"] = text_error
    method = ocr.status if ocr.text.strip() else ("pdf_text" if text.strip() else "no_text")
    return TextExtraction(
        combined,
        method,
        (1,) if ocr.text.strip() else tuple(pages),
        diagnostics,
        int((time.monotonic() - started) * 1000),
        False,
    )


def _process_worker(path_value: str, max_pages: int, output) -> None:
    try:
        output.put(("ok", _default_extract(Path(path_value), max_pages)))
    except Exception as error:
        output.put(("error", f"{type(error).__name__}: {error}"))


def _extract_with_timeout(
    path: Path,
    max_pages: int,
    timeout_seconds: int,
) -> TextExtraction:
    context = multiprocessing.get_context("spawn")
    output = context.Queue(maxsize=1)
    process = context.Process(
        target=_process_worker,
        args=(str(path), max_pages, output),
        daemon=True,
    )
    process.start()
    process.join(timeout_seconds)
    if process.is_alive():
        process.terminate()
        process.join(5)
        raise TimeoutError(f"OCR excedió {timeout_seconds} segundos")
    try:
        state, payload = output.get(timeout=1)
    except queue.Empty as error:
        raise RuntimeError("El extractor terminó sin devolver un resultado") from error
    finally:
        output.close()
    if state == "error":
        raise RuntimeError(str(payload))
    return payload


def get_document_text(
    connection: sqlite3.Connection | None,
    path: str | Path,
    *,
    extractor_version: str = TEXT_EXTRACTOR_VERSION,
    max_pages: int = 2,
    timeout_seconds: int = 45,
    extractor: Callable[[Path, int], TextExtraction] | None = None,
) -> TextExtraction:
    document_path = Path(path)
    sha256 = file_sha256(document_path)
    cached = _load_cache(connection, sha256, extractor_version)
    if cached is not None:
        return replace(cached, from_cache=True)

    try:
        result = (
            extractor(document_path, max_pages)
            if extractor is not None
            else _extract_with_timeout(document_path, max_pages, timeout_seconds)
        )
    except TimeoutError as error:
        result = TextExtraction(
            "",
            "timeout",
            (),
            {"code": "OCR_TIMEOUT", "detail": str(error)},
            0,
            False,
        )
    except (OSError, RuntimeError) as error:
        result = TextExtraction(
            "",
            "error",
            (),
            {"code": "TEXT_EXTRACTION_ERROR", "detail": str(error)},
            0,
            False,
        )

    result = replace(result, from_cache=False)
    _store_cache(connection, sha256, extractor_version, result)
    return result


__all__ = [
    "TEXT_EXTRACTOR_VERSION",
    "TextExtraction",
    "file_sha256",
    "get_document_text",
]
