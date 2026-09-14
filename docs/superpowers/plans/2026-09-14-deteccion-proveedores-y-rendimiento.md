# Detección de proveedores y rendimiento Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Clasificar las facturas de proveedores habituales sin intervención innecesaria y mantener ágil la revisión de expedientes con decenas de incidencias.

**Architecture:** `lector_pdf` conservará el texto y el motivo de una clasificación fallida; `source_analysis` decidirá si existe evidencia suficiente para reconocer una factura genérica. Los perfiles de proveedor se siguen declarando una sola vez en JSON. El lote compartirá ese catálogo y la interfaz sólo construirá una página de incidencias cada vez.

**Tech Stack:** Python 3, SQLite, CustomTkinter, pdfplumber, unittest.

**Spec:** `docs/superpowers/specs/2026-09-14-deteccion-proveedores-y-rendimiento-design.md`

## Global Constraints

- No aceptar como canónico un campo financiero, un consumo o un tipo de suministro sin evidencia inequívoca en la fuente.
- Mantener `config/proveedores.json` como catálogo global editable; los CUPS siguen asociados a la comunidad dentro del perfil.
- Conservar decisiones manuales y no reabrir sus incidencias durante un reanálisis.
- No incorporar PDFs privados de la comunidad al repositorio; las pruebas usan textos sintéticos representativos.
- El lote analiza fuera del hilo de interfaz y las incidencias muestran un máximo de diez tarjetas por página.

---

### Task 1: Reutilizar el catálogo de proveedores durante un lote

**Files:**
- Modify: `core/lector_pdf.py:37-52,888-935`
- Modify: `core/source_analysis.py:101-124,337-345`
- Test: `tests/test_source_analysis.py`

**Interfaces:**
- Produces: `lector_pdf.procesar_archivo(ruta_archivo, codigo_comunidad=None, con_bd=None, ruta_proveedores=None, proveedores=None) -> dict`.
- Produces: `source_analysis.analyse_pdf(path, pdf_processor=None, community_code=None, providers=None) -> SourceAnalysis` and `analyse_source(path, community_code, pdf_processor=None, providers=None) -> SourceAnalysis`.
- Consumes: the memoized `lector_pdf.cargar_proveedores(ruta_json)` result once per provider-config path and application process.

- [ ] **Step 1: Write the failing tests**

```python
def test_analyse_pdf_passes_a_preloaded_catalog_to_the_reader():
    providers = {"proveedores": {}}
    seen = {}
    def processor(_path, _community, *, proveedores, ruta_proveedores):
        seen["providers"] = proveedores
        return {"ok": True, "tipo": "FACTURA", "datos": {}}
    source_analysis.analyse_pdf(Path("invoice.pdf"), pdf_processor=processor,
                                community_code="658", providers=providers)
    self.assertIs(providers, seen["providers"])

def test_provider_catalog_is_loaded_once_per_path(self):
    lector_pdf.cargar_proveedores.cache_clear()
    config = str(PROJECT_ROOT / "config" / "proveedores.json")
    with mock.patch("lector_pdf.json.load", wraps=lector_pdf.json.load) as load:
        lector_pdf.cargar_proveedores(config)
        lector_pdf.cargar_proveedores(config)
    self.assertEqual(1, load.call_count)
    lector_pdf.cargar_proveedores.cache_clear()
```

- [ ] **Step 2: Run the focused tests to verify they fail**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_analysis -v`

Expected: FAIL because `analyse_pdf` does not accept `providers` and the case analyser has no shared catalogue.

- [ ] **Step 3: Implement the minimal catalogue hand-off**

```python
# lector_pdf.py
def procesar_archivo(ruta_archivo, codigo_comunidad=None, con_bd=None,
                     ruta_proveedores=None, proveedores=None):
    proveedores = proveedores if proveedores is not None else cargar_proveedores(ruta_proveedores)

# source_analysis.py
def analyse_pdf(path, *, pdf_processor=None, community_code=None, providers=None):
    result = processor(path, community_code, proveedores=providers,
                       ruta_proveedores=str(PROJECT_ROOT / "config" / "proveedores.json"))
```

Decorate `cargar_proveedores` with `functools.lru_cache` keyed by the explicit configuration path. Add the optional mapping hand-off to the reader and source analyser, while preserving the existing default path for callers that do not provide a catalogue. The cache is refreshed by restarting the local application after changing the JSON catalogue.

- [ ] **Step 4: Run focused tests to verify they pass**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_analysis tests.test_unified_ingestion_regressions -v`

Expected: PASS with the catalogue loaded once per batch and all prior ingestion regressions green.

- [ ] **Step 5: Commit**

```powershell
git add core/lector_pdf.py core/source_analysis.py core/case_ingestion.py core/expedient_ui.py tests/test_source_analysis.py
git commit -m "perf: share provider catalog across source batches"
```

### Task 2: Reconocer facturas genéricas con evidencia y contexto

**Files:**
- Modify: `core/lector_pdf.py:888-935`
- Modify: `core/source_analysis.py:31-64,101-124`
- Test: `tests/test_source_analysis.py`

**Interfaces:**
- Produces: `source_analysis.classify_generic_invoice_text(text: str) -> SourceAnalysis | None`.
- Produces: failed PDF reader results with `motivo`, `detalle` and a bounded `fragment` from the extracted text.
- Consumes: failed result dictionaries from `lector_pdf.procesar_archivo`.

- [ ] **Step 1: Write the failing tests**

```python
def test_unknown_provider_invoice_with_number_date_and_total_is_classified():
    result = source_analysis.analyse_pdf(
        Path("naturgy.pdf"),
        pdf_processor=lambda *_args, **_kwargs: {
            "ok": False, "motivo": "PROVEEDOR_NO_IDENTIFICADO",
            "detalle": "Ningún proveedor reconocido",
            "fragment": "Factura N.º FE26390022198715 Fecha de emisión: 08/06/2026 Total a pagar 326,98 €",
        },
    )
    self.assertEqual("invoice", result.kind)
    self.assertEqual("medium", result.confidence)
    self.assertEqual("326,98", result.candidates["importe_total"])

def test_unstructured_unknown_pdf_remains_classification_review():
    result = source_analysis.analyse_pdf(
        Path("nota.pdf"),
        pdf_processor=lambda *_args, **_kwargs: {"ok": False, "motivo": "PROVEEDOR_NO_IDENTIFICADO", "fragment": "Aviso interno"},
    )
    self.assertEqual("unknown", result.kind)
    self.assertIn("PROVEEDOR_NO_IDENTIFICADO", result.review_message)
```

- [ ] **Step 2: Run the focused tests to verify they fail**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_analysis.SourceAnalysisTest -v`

Expected: FAIL because unsuccessful reader results always become an undetailed `unknown` source.

- [ ] **Step 3: Implement the minimal safe generic classifier**

```python
_GENERIC_INVOICE_MARKERS = (r"\bfactura\b", r"n[.º°o]*\s*(?:de\s*)?factura")
_GENERIC_DATE_PATTERN = re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b")
_GENERIC_TOTAL_PATTERN = re.compile(
    r"(?:total\s+(?:a\s+pagar|factura|importe)|importe\s+total)\D{0,32}"
    r"(?P<valor>\d{1,3}(?:\.\d{3})*,\d{2}|\d+\.\d{2})",
    re.IGNORECASE,
)

def classify_generic_invoice_text(text: str) -> SourceAnalysis | None:
    normalized = " ".join(text.split())
    if not any(re.search(marker, normalized, re.IGNORECASE) for marker in _GENERIC_INVOICE_MARKERS):
        return None
    if not _GENERIC_DATE_PATTERN.search(normalized):
        return None
    total = _GENERIC_TOTAL_PATTERN.search(normalized)
    if total is None:
        return None
    return SourceAnalysis.invoice({"importe_total": total.group("valor")}, confidence="medium")
```

Add a maximum 1,000-character normalized fragment to failed reader responses after text extraction. Use the generic classifier only for `PROVEEDOR_NO_IDENTIFICADO`; return `unknown` for missing text, contradictory community/CUPS, or ambiguous content. Extend `SourceAnalysis.invoice` to accept `confidence="medium"` without changing the existing high-confidence default.

- [ ] **Step 4: Run focused tests to verify they pass**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_analysis tests.test_unified_ingestion_regressions -v`

Expected: PASS; generic invoice tests classify only supported evidence and existing reading behaviour remains unchanged.

- [ ] **Step 5: Commit**

```powershell
git add core/lector_pdf.py core/source_analysis.py tests/test_source_analysis.py
git commit -m "feat: classify evidenced invoices without provider profiles"
```

### Task 3: Añadir perfiles globales de los formatos observados en 658

**Files:**
- Modify: `config/proveedores.json`
- Test: `tests/test_source_analysis.py`

**Interfaces:**
- Produces: global profiles `MANTENIMIENTOS_ZARAGOZA` and `NATURGY_CLIENTES_GAS` with `firmas_identificacion`, `tipo_suministro`, extraction regexes and no invented community CUPS.
- Consumes: the existing `lector_pdf.identificar_proveedor` and `extraer_datos_factura` configuration schema.

- [ ] **Step 1: Write the failing tests**

```python
def test_catalogue_identifies_mantenimientos_zaragoza_invoice():
    providers = lector_pdf.cargar_proveedores(str(PROJECT_ROOT / "config" / "proveedores.json"))
    key, config = lector_pdf.identificar_proveedor(
        "INSTALACIONES ZARAGOZA S.L. MANTENIMIENTO DE SALAS DE CALDERAS",
        "F2524083.pdf", providers,
    )
    self.assertEqual("MANTENIMIENTOS_ZARAGOZA", key)
    self.assertEqual("MANTENIMIENTO", config["tipo_suministro"])

def test_catalogue_identifies_naturgy_clientes_gas_invoice():
    providers = lector_pdf.cargar_proveedores(str(PROJECT_ROOT / "config" / "proveedores.json"))
    key, config = lector_pdf.identificar_proveedor(
        "Naturgy Clientes, S.A.U. Estás en mercado libre. Período gas: del 26/04/2026 al 29/05/2026",
        "naturgy.pdf", providers,
    )
    self.assertEqual("NATURGY_CLIENTES_GAS", key)
    self.assertEqual("GAS", config["tipo_suministro"])
```

- [ ] **Step 2: Run the focused tests to verify they fail**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_analysis.SourceAnalysisTest -v`

Expected: FAIL because neither supplier profile exists in the catalogue.

- [ ] **Step 3: Add only evidence-based reusable profiles**

Add JSON entries using signatures seen in the text (`INSTALACIONES ZARAGOZA S.L.` and `Naturgy Clientes, S.A.U.`), an invoice number regex, labelled issue-date/period/total regexes, and supply type. Do not add a CUPS for 658 until the source's CUPS is explicitly confirmed; absent CUPS means `no_aplica`, not an inferred match.

- [ ] **Step 4: Run focused tests to verify they pass**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_analysis -v`

Expected: PASS; known patterns select their global profiles and no test depends on a private source file.

- [ ] **Step 5: Commit**

```powershell
git add config/proveedores.json tests/test_source_analysis.py
git commit -m "feat: add reusable provider profiles for common invoices"
```

### Task 4: Paginar y resumir incidencias en el panel principal

**Files:**
- Modify: `core/expedient_ui.py:119-180`
- Modify: `core/app.py:1168-1190,1315-1397`
- Test: `tests/test_expedient_ui.py`

**Interfaces:**
- Produces: `expedient_ui.issue_page(issues: Sequence[ReviewIssue], page: int, page_size: int = 10) -> tuple[tuple[ReviewIssue, ...], int, int]` returning items, normalized page and total pages.
- Consumes: `document_review.list_open_issues` output without changing database ordering.

- [ ] **Step 1: Write the failing tests**

```python
def test_issue_page_limits_first_view_to_ten_of_fifty_five(self):
    issues = tuple(range(55))
    visible, page, pages = expedient_ui.issue_page(issues, page=1)
    self.assertEqual(10, len(visible))
    self.assertEqual((1, 6), (page, pages))

def test_issue_page_normalizes_out_of_range_page(self):
    visible, page, pages = expedient_ui.issue_page(tuple(range(11)), page=9)
    self.assertEqual((2, 2), (page, pages))
    self.assertEqual(1, len(visible))
```

- [ ] **Step 2: Run the focused tests to verify they fail**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui -v`

Expected: FAIL because `issue_page` does not exist and the tray loops over every issue.

- [ ] **Step 3: Implement page state and compact controls**

```python
visible, current_page, total_pages = expedient_ui.issue_page(issues, self._issues_page)
ctk.CTkLabel(tray, text=f"{len(issues)} pendientes · Página {current_page}/{total_pages}")
ctk.CTkButton(controls, text="Anterior", command=lambda: self._set_issues_page(current_page - 1))
ctk.CTkButton(controls, text="Siguiente", command=lambda: self._set_issues_page(current_page + 1))
for issue in visible:
    render_issue_card(issue)
```

Initialize `_issues_page = 1`, reset it when the selected case changes, and use only `visible` when creating card widgets. Keep the existing open/resolve commands bound to the original `ReviewIssue` object.

- [ ] **Step 4: Run focused tests to verify they pass**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_ui tests.test_expedient_service -v`

Expected: PASS; 55 issues produce six pages and legacy UI tests remain green.

- [ ] **Step 5: Commit**

```powershell
git add core/expedient_ui.py core/app.py tests/test_expedient_ui.py
git commit -m "perf: paginate source review issues"
```

### Task 5: Verificar el flujo completo y documentar el reanálisis

**Files:**
- Modify: `GUIA_PROCESAR_TODO.txt`
- Test: `tests/test_unified_ingestion_regressions.py`

**Interfaces:**
- Consumes: the catalogue hand-off, generic classification and pagination from Tasks 1–4.
- Produces: an explicit user instruction to press `Reanalizar fuentes` after a catalogue update; automatic incidences are refreshed while manual corrections remain.

- [ ] **Step 1: Run the existing manual-outcome regression before documentation**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_unified_ingestion_regressions.UnifiedIngestionRegressionTest.test_legacy_204_generated_issues_are_reconciled_without_losing_manual_outcomes -v`

Expected: PASS; the existing regression proves `clear_open_automatic_issues` removes legacy automatic issues while preserving the manual correction and its resolution.

- [ ] **Step 2: Run the complete relevant verification**

Run: `& .\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_analysis tests.test_expedient_ui tests.test_expedient_service tests.test_unified_ingestion_regressions -v`

Expected: PASS with no failures.

- [ ] **Step 3: Update the user guide and commit**

Add: “Después de incorporar o modificar perfiles de proveedor, pulsa Reanalizar fuentes; no vuelvas a copiar los archivos.”

```powershell
git add GUIA_PROCESAR_TODO.txt
git commit -m "docs: explain source reanalysis after provider updates"
```
