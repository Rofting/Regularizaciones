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

`CURRENT_SCHEMA_VERSION` is now 7. New databases receive migration 6 after migrations 1–5 and migration 7 afterwards. Databases already at version 5 receive the two nullable analysis columns and the review-issue provenance column, each with a version record. The migrations' `PRAGMA table_info` guards make their column changes defensive.

## Self-review

- Confirmed reanalysis deletes only `open` `MISSING_REQUIRED_FIELD` and `DOCUMENT_CLASSIFICATION_REQUIRED` records.
- Confirmed analysis upserts refuse to replace rows whose `validation_status` is `validated`.
- Confirmed no write path targets `manual_corrections`, `facturas`, or canonical reading tables.
- Ran `git diff --check`; it reported no whitespace errors.

## Commit

`feat: persist analysed sources and safe reanalysis` (hash recorded in task delivery after commit).

## Concerns

None.

## Review-fix addendum

The four review findings were reproduced through public APIs before the fix.

- Resolved classifications now remain effective: a validated manual `document_kind` is used as the document's effective kind, and a resolved or dismissed classification outcome prevents a new blocking classification issue from being created.
- Migration 7 adds `review_issues.origin` (`automatic` or `manual`). Existing issues default to `manual` conservatively; only explicitly automatic open issues are deleted during reanalysis.
- Candidate replacement and obsolete-candidate cleanup now preserve both `validated` and `rejected` values.
- Analysed duplicate ingestion and reanalysis reload `SourceDocument` after persistence, so callers receive the current type and status.

The new regressions first failed against the reviewed implementation: manual same-code issues were deleted, rejected manual values became automatic candidates, and returned classification snapshots were stale. The focused verification then passed all 36 flow, review, and migration tests; full discovery also passed.

## Scoped re-review addendum

The scoped re-review identified that the persisted effective kind and the
generated requirements could diverge. The new regression starts with an
unknown source manually confirmed as `reading`, then supplies a conflicting
invoice analysis through both reanalysis and duplicate analysed ingestion.

The regression first failed because the invoice's three required fields were
used despite the effective reading classification. Requirement selection now
uses the effective document kind: an effective invoice has invoice fields,
while readings and other non-invoice kinds have none. The reading remains
persisted and returned as `reading` and has no invoice missing-field blockers.
The 37-test focused suite passed after the correction.
