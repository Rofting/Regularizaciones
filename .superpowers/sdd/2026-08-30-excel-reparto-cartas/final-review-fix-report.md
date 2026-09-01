# Final Review Fix Report

## Scope

Brief: `final-review-fix-brief.md`
Worktree: `feature/expedientes-incidencias`

This recovery started from an uncommitted worktree that already contained most
of the final-review implementation and tests. I preserved that work, added
missing boundary coverage for profile fingerprint checks, fixed a Windows test
cleanup issue, sanitized one remaining documented private path, and ran the
requested verification.

## RED / GREEN Evidence

### Profile fingerprint immutability

- Existing uncommitted implementation already compared `profile_sha256` in
  workflow, export, distribution, and letters.
- Added boundary tests for export, distribution, and letters using temporary
  copied profile JSON, then mutating the copied bytes while keeping key/version.
- RED limitation: these added tests passed immediately because the inherited
  implementation was already present when this recovery began.
- GREEN:
  `python -m unittest tests.test_excel_export_service.ExcelExportServiceTest.test_active_profile_rejects_same_version_with_changed_bytes tests.test_case_distribution.CaseDistributionTest.test_rejects_validated_export_when_profile_bytes_changed_without_version_bump tests.test_case_letter_service.CaseLetterServiceTest.test_rejects_letters_when_validated_profile_bytes_changed_without_version_bump -q`
  -> `Ran 3 tests ... OK`.

### Safe letter batch reuse

- Existing uncommitted tests cover changed selected concepts and changed
  historical graph input creating new batches.
- Existing implementation hashes canonical selected concepts, concept
  presentation metadata, `profile_sha256`, graph payload, identity metadata,
  and logo byte SHA when a logo exists.
- GREEN included in focused suite:
  `python -m unittest tests.test_case_letter_service -q` via the combined
  focused command below.

### Zero consumption and negative history

- Existing uncommitted tests cover equal readings as valid zero consumption and
  negative historical deltas omitted from graph history.
- Existing implementation allows zero owner consumption while still blocking
  negative counter resets, and filters negative historical rows instead of
  coercing them to zero.
- GREEN included in focused suite:
  `python -m unittest tests.test_case_distribution tests.test_case_letter_service -q`
  via the combined focused command below.

### Safe private validation resume

- Existing uncommitted tests cover unsafe resume paths, idempotent clean resume,
  open-issue blocking, and no duplicated outputs.
- RED observed during recovery:
  `python -m unittest tests.test_private_658_validation.Private658ValidationGuardsTest.test_resume_with_open_issues_stays_blocked_without_outputs -q`
  failed on Windows cleanup with `PermissionError` for the temporary
  `gestion.db`.
- Root cause: the new fixture used strict `TemporaryDirectory()` while SQLite
  cursors could keep the DB file handle alive briefly on Windows.
- GREEN after using the same
  `TemporaryDirectory(ignore_cleanup_errors=True)` pattern used by the existing
  file:
  combined focused suite below passed.

### Template installation after validation

- Existing uncommitted test covers invalid first workbook leaving no installed
  template/registration and valid retry succeeding with the same profile.
- Existing implementation defers installation/registration until after the
  bootstrap document has no open review issues.
- GREEN included in focused suite:
  `python -m unittest tests.test_excel_bootstrap_importer -q` via the combined
  focused command below.

### Private metadata guard

- Existing uncommitted guard scanned tracked `.md` and `.txt` files for known
  private identifiers and local user paths.
- I sanitized the remaining partially private workbook references in
  `docs/superpowers/specs/2026-08-30-generacion-excel-maestro-design.md`.
- I expanded the guard list to include the sanitized-away directory/file names.
- GREEN:
  `python -m unittest tests.test_private_658_validation.Private658ValidationGuardsTest.test_tracked_documents_do_not_contain_known_private_identifiers_or_local_paths -q`
  -> `Ran 1 test ... OK`.

## Verification

Focused final-review suite:

```text
<project-root>\.venv-fase1\Scripts\python.exe -m unittest tests.test_case_workflow_actions tests.test_excel_export_service tests.test_case_distribution tests.test_case_letter_service tests.test_private_658_validation tests.test_excel_bootstrap_importer -q
Ran 88 tests in 19.118s
OK
```

Full requested suite:

```text
<project-root>\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -t . -q
Ran 155 tests in 25.722s
OK
```

Compilation:

```text
<project-root>\.venv-fase1\Scripts\python.exe -m compileall -q core tests
exit 0
```

Whitespace check:

```text
git diff --check
exit 0
```

Private metadata scan:

```text
rg -n "658 estudio|estudio acs-cal|monasterio de poblet|pintor aguayo|reina fabiola|658 Regularizacion 2025 2026|Comunidad_644_LIQUIDADO|[A-Za-z]:\\Users\\" -g "*.md" -g "*.txt"
exit 1, no matches
```

## Limitations

- RED evidence for the inherited final-review behavior could not be reconstructed
  without undoing existing uncommitted work. The recovery preserves that work and
  records the observable RED/GREEN evidence from this session.
