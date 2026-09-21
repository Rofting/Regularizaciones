"""Bounded, deterministic batch analysis for source documents."""

from __future__ import annotations

import sqlite3
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

from document_text_service import file_sha256
from source_analysis import SourceAnalysis, analyse_source


@dataclass(frozen=True)
class BatchProgress:
    completed: int
    total: int
    filename: str
    phase: str


@dataclass(frozen=True)
class BatchItem:
    path: Path
    analysis: SourceAnalysis | None
    error: str | None


def analyse_batch(
    paths: Iterable[Path],
    *,
    community_code: str,
    database_path: str | Path,
    max_workers: int = 3,
    progress: Callable[[BatchProgress], None] | None = None,
    analyser: Callable[..., SourceAnalysis] = analyse_source,
) -> tuple[BatchItem, ...]:
    input_paths = tuple(Path(path) for path in paths)
    total = len(input_paths)
    emit = progress or (lambda _event: None)
    workers = min(3, max(1, int(max_workers)))

    sha_by_index: dict[int, str] = {}
    hash_errors: dict[int, str] = {}
    representatives: dict[str, Path] = {}
    for index, path in enumerate(input_paths):
        emit(BatchProgress(index, total, path.name, "hashing"))
        try:
            sha256 = file_sha256(path)
        except Exception as error:
            hash_errors[index] = f"{type(error).__name__}: {error}"
            continue
        sha_by_index[index] = sha256
        representatives.setdefault(sha256, path)

    def work(path: Path) -> BatchItem:
        connection = sqlite3.connect(database_path, timeout=5)
        try:
            connection.execute("PRAGMA busy_timeout=5000")
            for phase in ("text", "classification", "provider", "fields"):
                emit(BatchProgress(0, total, path.name, phase))
            analysis = analyser(
                path, community_code=community_code, connection=connection,
            )
            return BatchItem(path, analysis, None)
        except Exception as error:
            return BatchItem(path, None, f"{type(error).__name__}: {error}")
        finally:
            connection.close()

    by_sha: dict[str, BatchItem] = {}
    completed = 0
    with ThreadPoolExecutor(
        max_workers=workers, thread_name_prefix="source-analysis"
    ) as executor:
        futures = {
            executor.submit(work, path): sha256
            for sha256, path in representatives.items()
        }
        for future in as_completed(futures):
            sha256 = futures[future]
            item = future.result()
            by_sha[sha256] = item
            completed += 1
            emit(BatchProgress(completed, total, item.path.name, "completed"))

    results: list[BatchItem] = []
    for index, path in enumerate(input_paths):
        if index in hash_errors:
            results.append(BatchItem(path, None, hash_errors[index]))
            continue
        original = by_sha[sha_by_index[index]]
        results.append(BatchItem(path, original.analysis, original.error))

    emit(BatchProgress(total, total, "", "completed"))
    return tuple(results)


__all__ = ["BatchItem", "BatchProgress", "analyse_batch"]
