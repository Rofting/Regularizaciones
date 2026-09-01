# Final Review Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:test-driven-development to implement this plan task-by-task. This delivery is explicitly inline and must not create subagents.

**Goal:** Close the six final-review findings without inventing consumption or exposing private source metadata.

**Architecture:** Treat the configured profile bytes as immutable input shared by bootstrap, export, distribution, and letters. Make letter reuse depend on the complete rendered input, keep zero-consumption readings valid, add an idempotent resume boundary for isolated validation runs, and install Excel templates only after a clean bootstrap.

**Tech Stack:** Python 3, SQLite, unittest, openpyxl, python-docx.

**Spec:** `final-review-fix-brief.md`

## Global Constraints

- Do not read, copy, or version private data.
- Every behavior change starts with an observed failing test and ends with focused GREEN.
- Work only in the existing worktree and do not create subagents.
- Record RED/GREEN evidence in `.superpowers/sdd/2026-08-30-excel-reparto-cartas/final-review-fix-report.md`.

---

### Task 1: Immutable profile fingerprints

**Files:** `core/excel_profiles.py`, `core/excel_export_service.py`, `core/case_distribution.py`, `core/case_letter_service.py`, `core/case_workflow_actions.py`, focused tests.

- [ ] Add a regression test that mutates profile JSON while keeping key/version unchanged and expects export/distribution/letters to block.
- [ ] Run it and record the expected RED caused by key/version-only checks.
- [ ] Add a canonical profile-byte SHA helper, compare it with `excel_template_profiles.profile_sha256`, and include it in the canonical case input hash.
- [ ] Run focused profile/export/distribution/letter tests and record GREEN.

### Task 2: Complete letter reuse inputs

**Files:** `core/case_letter_service.py`, `tests/test_case_letter_service.py`.

- [ ] Add tests proving a different canonical concept selection and a changed graph/logo input create new batches.
- [ ] Run them and record RED reuse of the old batch.
- [ ] Hash selection, concept label/unit/order, verified profile SHA, graph history and logo byte SHA before reuse lookup.
- [ ] Run focused tests and record GREEN.

### Task 3: Zero and anomalous readings

**Files:** `core/case_distribution.py`, `core/case_letter_service.py`, distribution/letter tests.

- [ ] Add a test proving equal readings yield a valid zero weight and a test proving negative historical deltas are omitted.
- [ ] Run them and record RED.
- [ ] Accept zero consumption while retaining negative-reset blocking; filter negative history instead of coercing it to zero.
- [ ] Run focused tests and record GREEN.

### Task 4: Safe validation resume

**Files:** `core/private_658_validation.py`, `tests/test_private_658_validation.py`, `docs/validacion-excel-reparto-cartas.md`, `GUIA_PROCESAR_TODO.txt`.

- [ ] Add synthetic tests for resuming an existing safe run, blocking unsafe paths/open issues, and avoiding duplicate outputs.
- [ ] Run them and record RED for the absent resume entry point.
- [ ] Add a public resume function and `--resume` CLI that reuse `data/gestion.db`, validate the run boundary, and advance only the missing stages.
- [ ] Document the exact command and run focused GREEN.

### Task 5: Post-validation template installation

**Files:** `core/excel_bootstrap_importer.py`, `tests/test_excel_bootstrap_importer.py`.

- [ ] Add a test where an invalid first workbook leaves no template/registration and a valid second workbook with the same profile succeeds.
- [ ] Run it and record RED caused by early installation.
- [ ] Move installation/registration after clean parsing/validation while preserving the audited archived source.
- [ ] Run focused GREEN.

### Task 6: Private metadata guard and final verification

**Files:** tracked documentation, `tests/test_private_658_validation.py`, final report.

- [ ] Add a content scan test for tracked text documentation and observe RED on known private identifiers/local paths.
- [ ] Replace exact private names and local paths with generic markers and run GREEN.
- [ ] Run all focused tests, the full requested suite, compileall, and `git diff --check`.
- [ ] Write the evidence report, commit descriptive changes, and report remaining concerns.
