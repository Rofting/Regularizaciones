# Task 2 — Persist analysed sources and safe reanalysis

## Delivered

- Added schema migration 6. It adds nullable `source_documents.classification_confidence` and nullable `extraction_candidates.source_context`, checking existing columns before each `ALTER TABLE` so both fresh and already-migrated databases are safe.
- Added `add_analysed_document_to_case(...)`, consuming Task 1's `SourceAnalysis`. It records kind, confidence, candidates, and compact locator JSON. Analysis candidates begin as `candidate` values; they are not treated as user validation.
- A reading creates no invoice-required-field issues. An unknown source creates one idempotent open `DOCUMENT_CLASSIFICATION_REQUIRED` issue.
- Added `reanalyze_case_documents(...)`. It updates automatic analysis data, removes only open automatic missing-field/classification issues, then regenerates the applicable automatic review state. It preserves validated candidates, manual corrections, resolved/dismissed outcomes, and user-created issues.
- No canonical invoice or reading insertion was added.

## Red proof

Before implementation, ran:

```powershell
..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_flow tests.test_document_review -v
```

It failed as intended with six errors: the analysed ingestion API was absent and `record_candidates` did not yet accept analysis context.

## Green proof and tests

After implementation, the required focused suite passed:

```powershell
..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_flow tests.test_document_review -v
```

The expanded focused verification passed 32 tests:

```powershell
..\..\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_flow tests.test_document_review tests.test_db_migrations -v
```

The full test discovery was also run successfully:

```powershell
..\..\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -q
```

New focused coverage proves reading/unknown classification behavior, compact persistence metadata, manual candidate survival through reanalysis, automatic-issue-only cleanup, candidate context storage, and migration 6 columns/idempotency.

## Migration applied

`CURRENT_SCHEMA_VERSION` is now 6. New databases receive migration 6 after migrations 1–5; databases already at version 5 receive only the two nullable columns and a version-6 migration record. The migration's `PRAGMA table_info` guards make its column changes defensive.

## Self-review

- Confirmed reanalysis deletes only `open` `MISSING_REQUIRED_FIELD` and `DOCUMENT_CLASSIFICATION_REQUIRED` records.
- Confirmed analysis upserts refuse to replace rows whose `validation_status` is `validated`.
- Confirmed no write path targets `manual_corrections`, `facturas`, or canonical reading tables.
- Ran `git diff --check`; it reported no whitespace errors.

## Commit

`feat: persist analysed sources and safe reanalysis` (hash recorded in task delivery after commit).

## Concerns

None.
