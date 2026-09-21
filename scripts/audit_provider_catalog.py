"""Privacy-preserving audit of the global provider catalogue."""

from __future__ import annotations

import argparse
import json
import sqlite3
import sys
import time
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable, Iterable, Mapping


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from document_classifier import classify_document
from document_text_service import file_sha256, get_document_text
from provider_registry import ProviderProfile, load_provider_registry, resolve_provider


@dataclass(frozen=True)
class AuditReport:
    input_paths: int
    unique_documents: int
    recognised_invoices: int
    provider_decisions_remaining: int
    non_invoice_false_positives: int
    document_types: Mapping[str, int]
    providers: Mapping[str, int]
    confidences: Mapping[str, int]
    diagnostic_codes: Mapping[str, int]
    total_duration_ms: int

    @property
    def text_invoices_recognised_pct(self) -> float:
        invoice_total = self.recognised_invoices + self.provider_decisions_remaining
        return 100.0 if invoice_total == 0 else round(
            (self.recognised_invoices / invoice_total) * 100.0, 2
        )


def audit_documents(
    paths: Iterable[Path],
    *,
    registry: Mapping[str, ProviderProfile],
    text_loader: Callable[[Path], str],
) -> AuditReport:
    path_list = tuple(Path(path) for path in paths)
    seen: set[str] = set()
    kinds: Counter[str] = Counter()
    providers: Counter[str] = Counter()
    confidences: Counter[str] = Counter()
    diagnostics: Counter[str] = Counter()
    recognised = 0
    unresolved = 0
    false_positives = 0
    started = time.monotonic()

    for path in path_list:
        try:
            sha256 = file_sha256(path)
        except OSError:
            diagnostics["FILE_UNREADABLE"] += 1
            continue
        if sha256 in seen:
            continue
        seen.add(sha256)
        try:
            text = text_loader(path)
            classification = classify_document(text, path.name)
            kinds[classification.kind] += 1
            confidences[classification.confidence] += 1
            if classification.kind not in {"invoice", "credit_note"}:
                continue
            match = resolve_provider(
                registry, text, path.name, classification.kind
            )
            if match is None:
                unresolved += 1
                diagnostics["PROVIDER_UNRESOLVED"] += 1
            else:
                recognised += 1
                providers[match.provider_key] += 1
        except Exception as error:
            diagnostics[type(error).__name__.upper()] += 1

    return AuditReport(
        input_paths=len(path_list),
        unique_documents=len(seen),
        recognised_invoices=recognised,
        provider_decisions_remaining=unresolved,
        non_invoice_false_positives=false_positives,
        document_types=dict(sorted(kinds.items())),
        providers=dict(sorted(providers.items())),
        confidences=dict(sorted(confidences.items())),
        diagnostic_codes=dict(sorted(diagnostics.items())),
        total_duration_ms=int((time.monotonic() - started) * 1000),
    )


def _main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", required=True, type=Path)
    parser.add_argument("--database", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument(
        "--catalog",
        type=Path,
        default=PROJECT_ROOT / "config" / "proveedores.json",
    )
    args = parser.parse_args()

    paths = tuple(sorted(args.root.rglob("*.pdf")))
    registry = load_provider_registry(args.catalog)
    connection = sqlite3.connect(args.database)
    try:
        report = audit_documents(
            paths,
            registry=registry,
            text_loader=lambda path: get_document_text(connection, path).text,
        )
    finally:
        connection.close()

    args.output.parent.mkdir(parents=True, exist_ok=True)
    payload = asdict(report)
    payload["text_invoices_recognised_pct"] = report.text_invoices_recognised_pct
    args.output.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"unique_pdf={report.unique_documents}")
    print(f"text_invoices_recognised_pct={report.text_invoices_recognised_pct}")
    print(f"provider_decisions_remaining={report.provider_decisions_remaining}")
    print(f"non_invoice_false_positives={report.non_invoice_false_positives}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
