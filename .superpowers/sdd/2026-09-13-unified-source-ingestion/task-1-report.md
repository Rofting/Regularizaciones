# Task 1 report — source analysis model

## Tests run

1. Red: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_source_analysis -v`
   failed as intended with `ModuleNotFoundError: No module named 'source_analysis'` before the module existed.
2. Green: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_source_analysis -v`
   passed: 3 tests, 0 failures.
3. `git diff --check` completed without whitespace errors.

## Files changed

- `core/source_analysis.py`: immutable analysis/location contracts, PDF dispatch and mapping, tabular-header classification, and CSV/XLS/XLSX header readers.
- `tests/test_source_analysis.py`: focused invoice, reading, and unknown-classification regression tests.

## Commit

`feat: classify sources before review`

## Self-review

- `SourceAnalysis` is frozen and snapshots candidates into a read-only mapping; required fields are tuples.
- Invoice requirements are exactly `fecha_inicio`, `fecha_fin`, and `importe_total`; reading results have none.
- Values extracted from PDFs are represented as strings or `None`; uncertain values are never replaced with zero.
- PDF page/fragment and tabular sheet/cell data are retained whenever supplied by the reader/parser.
- Header matching is accent-insensitive and does not contain community-specific rules.

## Concerns

None.

## Review-fix addendum

The Task 1 review identified that `Propiedad` and `Vivienda` headers were
incorrectly returned as `unknown`. A regression test now covers both headers,
and the owner-list signal treats `propiedad`, `vivienda`, and `propietario` as
owner indicators while retaining reading-header precedence.

Tests run:

1. Red: `..\\..\\.venv-fase1\\Scripts\\python.exe -m unittest tests.test_source_analysis -v`
   failed for `Propiedad` and `Vivienda`, each returning `unknown` instead of `owners`.
2. Green: the same command passed: 4 tests, 0 failures.
3. `git diff --check` completed without whitespace errors.
