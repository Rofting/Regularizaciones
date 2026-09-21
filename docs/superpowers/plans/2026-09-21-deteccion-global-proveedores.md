# Detección global fiable de proveedores — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convertir el análisis de PDF en una canalización global, trazable y rápida que reconozca tipo documental, proveedor y campos con confianza suficiente, y que solo aplique al reparto los documentos elegibles para la comunidad y el período.

**Architecture:** Se separan cinco responsabilidades: obtención/caché de texto, clasificación documental, identidad global del proveedor, extracción por familias y elegibilidad del expediente. `source_analysis.py` orquesta esas piezas y conserva un adaptador temporal al lector actual; `case_ingestion.py` persiste evidencia y evita que un documento reconocido pero no aplicable alcance las tablas canónicas.

**Tech Stack:** Python 3, SQLite, `pdfplumber`, `pdf2image`, RapidOCR, Tesseract opcional, `unittest`, CustomTkinter y JSON versionado.

**Spec:** `docs/superpowers/specs/2026-09-21-deteccion-global-proveedores-design.md`

## Global Constraints

- No enviar PDF, texto extraído ni datos personales a servicios externos.
- No inventar un valor: confianza baja nunca entra en facturas, lecturas, reparto ni cartas.
- Mantener compatibles los 16 perfiles actuales mientras se migra el catálogo.
- Una factura reconocida solo se aplica si coinciden comunidad, período y concepto/módulo activo.
- Cachear texto por `(sha256, extractor_version)` y limitar cada análisis a dos páginas en la primera pasada.
- Un archivo bloqueado o ilegible no detiene el lote.
- No añadir documentos privados al repositorio; las pruebas versionadas usan fragmentos sintéticos.
- No añadir ni restaurar `data/gestion.db` en ningún commit.
- Antes de cada commit, ejecutar la prueba indicada y `git diff --check` sobre los archivos de la tarea.

---

### Task 0: Preservar la línea base local ya verificada

**Files:**
- Modify/commit existing: `config/proveedores.json`
- Modify/commit existing: `core/case_ingestion.py`
- Modify/commit existing: `core/community_discovery.py`
- Modify/commit existing: `core/expedient_ui.py`
- Modify/commit existing: `core/lector_pdf.py`
- Modify/commit existing: `core/source_analysis.py`
- Modify/commit existing: `tests/test_expedient_flow.py`
- Modify/commit existing: `tests/test_source_analysis.py`
- Exclude: `data/gestion.db`

**Interfaces:**
- Consumes: estado local que ya reduce incidencias y conserva reanálisis seguro.
- Produces: un commit limpio que sirve como base para las tareas siguientes.

- [ ] **Step 1: Verificar exactamente la línea base**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -v
```

Expected: `428` pruebas o más, `OK`, sin errores ni fallos.

- [ ] **Step 2: Revisar que la base de datos no se vaya a incluir**

Run:

```powershell
git status --short
git diff --check -- config/proveedores.json core/case_ingestion.py core/community_discovery.py core/expedient_ui.py core/lector_pdf.py core/source_analysis.py tests/test_expedient_flow.py tests/test_source_analysis.py
```

Expected: `data/gestion.db` puede aparecer modificada, pero no hay errores de espacios en los ocho archivos de código/configuración.

- [ ] **Step 3: Crear el commit de línea base con selección explícita**

```powershell
git add -- config/proveedores.json core/case_ingestion.py core/community_discovery.py core/expedient_ui.py core/lector_pdf.py core/source_analysis.py tests/test_expedient_flow.py tests/test_source_analysis.py
git diff --cached --name-only
git commit -m "fix: consolidar deteccion actual de fuentes"
```

Expected: la lista staged contiene únicamente esos ocho archivos y nunca `data/gestion.db`.

---

### Task 1: Persistencia de caché, evidencia y elegibilidad

**Files:**
- Modify: `core/db_migrations.py:735-789`
- Modify: `tests/test_db_migrations.py:13-103`

**Interfaces:**
- Consumes: `db_migrations.migrate(connection) -> int` y migración vigente `13`.
- Produces: migración `14`; tablas `document_text_cache`, `source_field_evidence`; columnas `provider_key`, `analysis_version`, `eligibility_status`, `eligibility_reason` en `source_documents`.

- [ ] **Step 1: Escribir la prueba de migración 14**

Añadir a `DatabaseMigrationTest`:

```python
def test_migration_fourteen_adds_detection_cache_evidence_and_eligibility(self):
    with redirect_stdout(StringIO()):
        gestor_bd.crear_bd(str(self.database_path))

    with closing(self._connect()) as connection:
        versions = [row[0] for row in connection.execute(
            "SELECT version FROM schema_migrations ORDER BY version"
        )]
        tables = {row[0] for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        )}
        document_columns = {row[1] for row in connection.execute(
            "PRAGMA table_info(source_documents)"
        )}

    self.assertEqual(list(range(1, 15)), versions)
    self.assertTrue({"document_text_cache", "source_field_evidence"}.issubset(tables))
    self.assertTrue({
        "provider_key", "analysis_version", "eligibility_status", "eligibility_reason",
    }.issubset(document_columns))
```

- [ ] **Step 2: Ejecutar la prueba y comprobar que falla**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest tests.test_db_migrations.DatabaseMigrationTest.test_migration_fourteen_adds_detection_cache_evidence_and_eligibility -v
```

Expected: FAIL porque la versión máxima todavía es `13`.

- [ ] **Step 3: Implementar la migración 14**

Añadir a `core/db_migrations.py`:

```python
def _migration_14(connection: sqlite3.Connection) -> None:
    columns = {row[1] for row in connection.execute(
        "PRAGMA table_info(source_documents)"
    )}
    additions = {
        "provider_key": "TEXT",
        "analysis_version": "TEXT",
        "eligibility_status": "TEXT NOT NULL DEFAULT 'pending'",
        "eligibility_reason": "TEXT",
    }
    for name, declaration in additions.items():
        if name not in columns:
            connection.execute(
                f"ALTER TABLE source_documents ADD COLUMN {name} {declaration}"
            )
    connection.execute("""CREATE TABLE IF NOT EXISTS document_text_cache (
        sha256 TEXT NOT NULL,
        extractor_version TEXT NOT NULL,
        text_content TEXT NOT NULL,
        method TEXT NOT NULL,
        pages_json TEXT NOT NULL,
        diagnostics_json TEXT NOT NULL,
        duration_ms INTEGER NOT NULL,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        PRIMARY KEY (sha256, extractor_version)
    )""")
    connection.execute("""CREATE TABLE IF NOT EXISTS source_field_evidence (
        id_evidence INTEGER PRIMARY KEY AUTOINCREMENT,
        id_document INTEGER NOT NULL REFERENCES source_documents(id_document) ON DELETE CASCADE,
        field_name TEXT NOT NULL,
        value TEXT,
        confidence TEXT NOT NULL CHECK(confidence IN ('high','medium','low')),
        source TEXT NOT NULL,
        locator_json TEXT NOT NULL DEFAULT '{}',
        rule_id TEXT NOT NULL,
        extractor_version TEXT NOT NULL,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        UNIQUE(id_document, field_name, rule_id, extractor_version)
    )""")
```

Registrar `14: _migration_14` en `MIGRATIONS` y actualizar las expectativas de versiones existentes a `range(1, 15)`.

- [ ] **Step 4: Ejecutar las pruebas de migración**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest tests.test_db_migrations -v
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add core/db_migrations.py tests/test_db_migrations.py
git commit -m "feat: persistir cache y evidencia documental"
```

---

### Task 2: Clasificador documental anterior al proveedor

**Files:**
- Create: `core/document_classifier.py`
- Create: `tests/test_document_classifier.py`

**Interfaces:**
- Consumes: texto normalizado y nombre de archivo.
- Produces: `DocumentClassification(kind: str, confidence: str, rule_id: str, evidence: tuple[str, ...])` y `classify_document(text: str, filename: str) -> DocumentClassification`.

- [ ] **Step 1: Escribir pruebas para tipos operativos y exclusiones**

```python
from document_classifier import classify_document


class DocumentClassifierTest(unittest.TestCase):
    def test_invoice_requires_invoice_structure_not_only_an_amount(self):
        result = classify_document(
            "FACTURA Nº F-102 Fecha 12/03/2026 Base imponible 100,00 IVA 21,00 Total 121,00 €",
            "F-102.pdf",
        )
        self.assertEqual(("invoice", "high"), (result.kind, result.confidence))

    def test_quote_wins_over_generic_total(self):
        result = classify_document(
            "PRESUPUESTO Nº P-8 Fecha 12/03/2026 Total 121,00 € Validez 30 días",
            "presupuesto.pdf",
        )
        self.assertEqual("quote", result.kind)

    def test_delivery_note_and_bank_receipt_never_become_invoices(self):
        fixtures = (
            ("ALBARÁN Entrega de gasóleo 900 litros", "delivery_note"),
            ("JUSTIFICANTE DE TRANSFERENCIA IBAN ES00 Ordenante Comunidad", "bank_receipt"),
        )
        for text, expected in fixtures:
            with self.subTest(expected=expected):
                self.assertEqual(expected, classify_document(text, "documento.pdf").kind)

    def test_meter_table_is_reading(self):
        result = classify_document(
            "Propiedad Lectura anterior Lectura actual Consumo ACS PA2-1A 144 166 22",
            "Lecturas consumo 07 2025 a 07 2026.pdf",
        )
        self.assertEqual("reading", result.kind)
```

- [ ] **Step 2: Ejecutar las pruebas y comprobar que fallan**

Run: `.\.venv-fase1\Scripts\python.exe -m unittest tests.test_document_classifier -v`

Expected: ERROR `ModuleNotFoundError: document_classifier`.

- [ ] **Step 3: Implementar reglas ordenadas con exclusiones prioritarias**

```python
@dataclass(frozen=True)
class DocumentClassification:
    kind: str
    confidence: str
    rule_id: str
    evidence: tuple[str, ...]


def classify_document(text: str, filename: str) -> DocumentClassification:
    normalized = normalize_detection_text(f"{filename}\n{text}")
    for kind, markers in EXCLUSION_RULES:
        matched = tuple(marker for marker in markers if marker.search(normalized))
        if matched:
            return DocumentClassification(kind, "high", f"document:{kind}:v1",
                                          tuple(item.pattern for item in matched))
    if _looks_like_reading(normalized):
        return DocumentClassification("reading", "high", "document:reading:v1", ("meter-table",))
    if _looks_like_invoice(normalized):
        return DocumentClassification("invoice", "high", "document:invoice:v1", ("invoice-structure",))
    return DocumentClassification("unknown", "low", "document:unknown:v1", ())
```

Definir reglas explícitas para `credit_note`, `owners`, `quote`, `delivery_note`, `bank_receipt`, `report`, `other` y `unknown`. Las exclusiones se evalúan antes que la presencia de total o fecha.

- [ ] **Step 4: Ejecutar pruebas**

Run: `.\.venv-fase1\Scripts\python.exe -m unittest tests.test_document_classifier -v`

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add core/document_classifier.py tests/test_document_classifier.py
git commit -m "feat: clasificar documentos antes del proveedor"
```

---

### Task 3: Registro global y resolución fiscal de proveedores

**Files:**
- Create: `core/provider_registry.py`
- Create: `tests/test_provider_registry.py`
- Modify: `config/proveedores.json`
- Modify: `tests/test_source_analysis.py`

**Interfaces:**
- Consumes: catálogo JSON compatible, `DocumentClassification`, texto y nombre.
- Produces: `ProviderProfile`, `ProviderMatch`; `provider_registry_from_payload(payload) -> Mapping[str, ProviderProfile]`; `load_provider_registry(path) -> Mapping[str, ProviderProfile]`; `resolve_provider(registry, text, filename, document_kind) -> ProviderMatch | None`.

- [ ] **Step 1: Escribir pruebas de identidad, colisión y compatibilidad**

```python
def _profile(*, aliases, document_types=("invoice",)):
    return {
        "tax_ids": [], "aliases": aliases,
        "document_types": list(document_types),
        "service_family": "MANTENIMIENTO",
        "extractor_family": "standard_spanish_invoice",
        "required_signatures": [r"\bFACTURA\b"],
        "excluded_signatures": [r"\bPRESUPUESTO\b"],
    }


def test_tax_id_beats_customer_and_bank_names(self):
    registry = provider_registry_from_payload({"proveedores": {
        "EMISOR": {"tax_ids": ["B12345678"], "aliases": ["Emisor Real"],
                   "document_types": ["invoice"], "service_family": "MANTENIMIENTO",
                   "extractor_family": "standard_spanish_invoice"},
        "BANCO": {"tax_ids": ["A87654321"], "aliases": ["Banco Ejemplo"],
                  "document_types": ["bank_receipt"], "service_family": "PAGO",
                  "extractor_family": "standard_spanish_invoice"},
    }})
    match = resolve_provider(
        registry,
        "Cliente Comunidad CIF H00000000 Banco Ejemplo Emisor Real CIF B12345678 FACTURA F-9",
        "factura.pdf", "invoice",
    )
    self.assertEqual(("EMISOR", "high"), (match.provider_key, match.confidence))

def test_equal_alias_evidence_is_ambiguous(self):
    registry = provider_registry_from_payload({"proveedores": {
        "UNO": _profile(aliases=["ACME"]),
        "DOS": _profile(aliases=["ACME"]),
    }})
    self.assertIsNone(resolve_provider(registry, "ACME FACTURA F-1", "f.pdf", "invoice"))

def test_provider_document_type_must_match(self):
    registry = provider_registry_from_payload({"proveedores": {
        "UNO": _profile(aliases=["ACME"], document_types=["invoice"]),
    }})
    self.assertIsNone(resolve_provider(registry, "ACME PRESUPUESTO", "p.pdf", "quote"))
```

- [ ] **Step 2: Ejecutar y comprobar fallo**

Run: `.\.venv-fase1\Scripts\python.exe -m unittest tests.test_provider_registry -v`

Expected: ERROR porque el módulo aún no existe.

- [ ] **Step 3: Implementar carga compatible y validación estricta**

```python
@dataclass(frozen=True)
class ProviderProfile:
    key: str
    display_name: str
    tax_ids: tuple[str, ...]
    aliases: tuple[str, ...]
    document_types: tuple[str, ...]
    service_family: str
    extractor_family: str
    required_signatures: tuple[str, ...]
    excluded_signatures: tuple[str, ...]
    legacy: Mapping[str, object]


@dataclass(frozen=True)
class ProviderMatch:
    provider_key: str
    confidence: str
    rule_id: str
    evidence: tuple[str, ...]
```

Implementar `provider_registry_from_payload` como núcleo validable y hacer que `load_provider_registry` se limite a leer JSON y delegar en él. Normalizar CIF/NIF eliminando separadores, validar letra/número estructural y rechazar identificadores duplicados. Para perfiles antiguos, traducir `nombre_display`, `tipo_suministro`, `firmas_identificacion`, `firmas_requeridas` y `firma_exclusion` sin alterar sus regex.

- [ ] **Step 4: Ampliar el esquema del JSON sin romper los 16 perfiles existentes**

Añadir `_schema_version: "2"` y, de forma incremental, `tax_ids`, `aliases`, `document_types`, `service_family`, `extractor_family`, `required_signatures`, `excluded_signatures`. Un perfil antiguo sin esos campos se adapta en memoria; el validador sigue rechazando colisiones nuevas.

- [ ] **Step 5: Ejecutar pruebas del registro y regresión actual**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest tests.test_provider_registry tests.test_source_analysis -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add core/provider_registry.py config/proveedores.json tests/test_provider_registry.py tests/test_source_analysis.py
git commit -m "feat: resolver proveedores por identidad global"
```

---

### Task 4: Servicio único de texto y OCR cacheado

**Files:**
- Create: `core/document_text_service.py`
- Create: `tests/test_document_text_service.py`
- Modify: `core/lector_pdf.py:389-752`

**Interfaces:**
- Consumes: ruta PDF, conexión SQLite opcional y extractores actuales de `lector_pdf`.
- Produces: `TextExtraction`; `get_document_text(connection, path, *, extractor_version="document-text-v1", max_pages=2, timeout_seconds=45, extractor=None) -> TextExtraction`.

- [ ] **Step 1: Escribir pruebas de caché, OCR selectivo y recuperación**

```python
class DocumentTextServiceTest(unittest.TestCase):
    def test_same_sha_and_version_uses_cached_text(self):
        calls = []
        def extractor(path, max_pages):
            calls.append(path)
            return TextExtraction("texto factura", "pdf_text", (1,), {}, 3, False)

        first = get_document_text(self.connection, self.path, extractor=extractor)
        second = get_document_text(self.connection, self.path, extractor=extractor)

        self.assertEqual("texto factura", second.text)
        self.assertEqual(1, len(calls))
        self.assertFalse(first.from_cache)
        self.assertTrue(second.from_cache)

    def test_new_extractor_version_invalidates_cache(self):
        calls = []
        extractor = lambda path, max_pages: (
            calls.append(path) or TextExtraction("texto", "pdf_text", (1,), {}, 1, False)
        )
        get_document_text(self.connection, self.path, extractor_version="v1", extractor=extractor)
        get_document_text(self.connection, self.path, extractor_version="v2", extractor=extractor)
        self.assertEqual(2, len(calls))

    def test_timeout_returns_recoverable_diagnostic(self):
        def timed_out(_path, _max_pages):
            raise TimeoutError("OCR excedió 45 segundos")
        result = get_document_text(self.connection, self.path, extractor=timed_out)
        self.assertEqual("timeout", result.method)
        self.assertEqual("OCR_TIMEOUT", result.diagnostics["code"])
```

- [ ] **Step 2: Ejecutar y comprobar fallo**

Run: `.\.venv-fase1\Scripts\python.exe -m unittest tests.test_document_text_service -v`

Expected: ERROR porque el servicio aún no existe.

- [ ] **Step 3: Implementar resultado, huella y caché**

```python
@dataclass(frozen=True)
class TextExtraction:
    text: str
    method: str
    pages: tuple[int, ...]
    diagnostics: Mapping[str, object]
    duration_ms: int
    from_cache: bool


def get_document_text(connection, path: Path, *,
                      extractor_version: str = TEXT_EXTRACTOR_VERSION,
                      max_pages: int = 2, timeout_seconds: int = 45,
                      extractor=None) -> TextExtraction:
    sha256 = file_sha256(path)
    cached = _load_cache(connection, sha256, extractor_version)
    if cached is not None:
        return replace(cached, from_cache=True)
    try:
        if extractor is not None:
            result = extractor(path, max_pages)
        else:
            result = _extract_with_timeout(path, max_pages, timeout_seconds)
    except TimeoutError as error:
        result = TextExtraction("", "timeout", (),
                                {"code": "OCR_TIMEOUT", "detail": str(error)}, 0, False)
    _store_cache(connection, sha256, extractor_version, result)
    return result
```

La extracción normal usa capa PDF primero. Solo activa RapidOCR/Tesseract cuando el texto es insuficiente para señales esenciales. El worker de producción corre en un proceso con `spawn`; al superar 45 segundos se termina, devuelve `OCR_TIMEOUT` y permite continuar con el siguiente documento.

- [ ] **Step 4: Convertir `lector_pdf` en adaptador de compatibilidad**

Mantener `extraer_texto`, `extraer_texto_ocr_con_diagnostico` y `procesar_archivo`, pero permitir que `procesar_archivo(..., extracted_text=None)` reutilice el texto entregado por el servicio y no vuelva a abrir el PDF.

- [ ] **Step 5: Ejecutar pruebas OCR y del servicio**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest tests.test_document_text_service tests.test_source_analysis.SourceAnalysisTest.test_ocr_limits_scanned_pdf_to_the_invoice_cover_page tests.test_source_analysis.SourceAnalysisTest.test_image_only_pdf_is_routed_to_ocr_before_pdfplumber -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add core/document_text_service.py core/lector_pdf.py tests/test_document_text_service.py
git commit -m "feat: cachear texto y limitar OCR por documento"
```

---

### Task 5: Extractores por familia y confianza por campo

**Files:**
- Create: `core/invoice_extractors.py`
- Create: `tests/test_invoice_extractors.py`
- Modify: `core/source_analysis.py:24-200`

**Interfaces:**
- Consumes: `ProviderProfile` y texto ya extraído.
- Produces: `FieldEvidence`; `ExtractionBundle`; `extract_invoice_fields(profile, text) -> ExtractionBundle`.

- [ ] **Step 1: Escribir pruebas de fechas, importes y reconciliación**

```python
def profile(extractor_family: str) -> ProviderProfile:
    return ProviderProfile(
        key="SYNTHETIC", display_name="Proveedor sintético", tax_ids=(),
        aliases=("PROVEEDOR SINTETICO",), document_types=("invoice",),
        service_family="MANTENIMIENTO", extractor_family=extractor_family,
        required_signatures=(), excluded_signatures=(), legacy={},
    )


def test_standard_invoice_returns_field_level_evidence(self):
    result = extract_invoice_fields(
        profile("standard_spanish_invoice"),
        "Factura F-9 Fecha 01/04/2026 Periodo 01/03/2026 a 31/03/2026 "
        "Base imponible 100,00 € IVA 21,00 € Total 121,00 €",
    )
    self.assertEqual("2026-03-01", result.fields["fecha_inicio"].value)
    self.assertEqual("2026-03-31", result.fields["fecha_fin"].value)
    self.assertEqual("121.00", result.fields["importe_total"].value)
    self.assertEqual("high", result.fields["importe_total"].confidence)

def test_total_that_does_not_reconcile_is_not_high_confidence(self):
    result = extract_invoice_fields(
        profile("standard_spanish_invoice"),
        "Factura F-9 Base imponible 100,00 € IVA 21,00 € Total 999,00 €",
    )
    self.assertNotEqual("high", result.fields["importe_total"].confidence)

def test_unlabelled_numbers_are_never_dates_or_totals(self):
    result = extract_invoice_fields(profile("standard_spanish_invoice"), "1231 1233 131231")
    self.assertNotIn("fecha_inicio", result.fields)
    self.assertNotIn("importe_total", result.fields)
```

- [ ] **Step 2: Ejecutar y comprobar fallo**

Run: `.\.venv-fase1\Scripts\python.exe -m unittest tests.test_invoice_extractors -v`

Expected: ERROR porque el módulo aún no existe.

- [ ] **Step 3: Implementar modelos y familias**

```python
@dataclass(frozen=True)
class FieldEvidence:
    value: str
    confidence: str
    source: str
    locator: Mapping[str, object]
    rule_id: str
    extractor_version: str


@dataclass(frozen=True)
class ExtractionBundle:
    fields: Mapping[str, FieldEvidence]
    diagnostics: tuple[str, ...] = ()
```

Registrar extractores para `standard_spanish_invoice`, `electricity`, `gas_fuel`, `periodic_maintenance`, `elevators`, `boilers_hvac`, `metering_management` y `water_public_fees`. Cada familia reutiliza parseadores comunes de importe, fecha, período e IVA; los formatos del proveedor solo reemplazan reglas concretas.

- [ ] **Step 4: Integrar evidencia en `SourceAnalysis` sin romper llamadas actuales**

Extender el dataclass con valores por defecto:

```python
provider_key: str | None = None
field_evidence: Mapping[str, FieldEvidence] = field(default_factory=dict)
analysis_version: str = "source-analysis-v2"
```

`candidates` sigue exponiendo `{campo: evidencia.value}` para compatibilidad con `case_ingestion.py`.

- [ ] **Step 5: Ejecutar pruebas**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest tests.test_invoice_extractors tests.test_source_analysis -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add core/invoice_extractors.py core/source_analysis.py tests/test_invoice_extractors.py tests/test_source_analysis.py
git commit -m "feat: extraer facturas por familia con confianza"
```

---

### Task 6: Orquestación, elegibilidad y una sola incidencia útil

**Files:**
- Create: `core/source_eligibility.py`
- Create: `tests/test_source_eligibility.py`
- Modify: `core/source_analysis.py:150-200,582-595`
- Modify: `core/case_ingestion.py:151-363,483-552`
- Modify: `core/document_review.py:298-550`
- Modify: `tests/test_expedient_flow.py`
- Modify: `tests/test_unified_ingestion_regressions.py`

**Interfaces:**
- Consumes: tipo documental, proveedor, `ExtractionBundle`, perfil Excel activo y fechas del expediente.
- Produces: `EligibilityDecision(status, reason, concept_keys)`; persistencia de evidencia; una incidencia agrupada cuando la decisión raíz no es segura.

- [ ] **Step 1: Escribir pruebas de elegibilidad**

```python
def test_recognised_invoice_outside_active_modules_is_archived_not_applied(self):
    decision = evaluate_invoice_eligibility(
        community_code="658",
        case_start=date(2025, 9, 1),
        case_end=date(2026, 8, 31),
        active_modules=("GAS", "ELECTRICIDAD", "AGUA", "ACS"),
        document_kind="invoice",
        service_family="ASCENSORES",
        period_start=date(2026, 4, 1),
        period_end=date(2026, 4, 30),
        community_confidence="high",
    )
    self.assertEqual(("not_applicable", "service_not_active"),
                     (decision.status, decision.reason))

def test_invoice_for_another_period_requires_one_grouped_decision(self):
    decision = evaluate_invoice_eligibility(
        community_code="658", case_start=date(2025, 9, 1), case_end=date(2026, 8, 31),
        active_modules=("GAS",), document_kind="invoice", service_family="GAS",
        period_start=date(2024, 1, 1), period_end=date(2024, 1, 31),
        community_confidence="high",
    )
    self.assertEqual(("review_required", "period_outside_case"),
                     (decision.status, decision.reason))
```

- [ ] **Step 2: Ejecutar y comprobar fallo**

Run: `.\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_eligibility -v`

Expected: ERROR porque el módulo aún no existe.

- [ ] **Step 3: Implementar decisión pura y mapeo de servicios**

```python
SERVICE_MODULES = {
    "ELECTRICIDAD": ("ELECTRICIDAD",),
    "GAS": ("GAS",),
    "COMBUSTIBLE": ("GAS",),
    "AGUA": ("AGUA",),
    "MANTENIMIENTO": ("OTROS_GASTOS",),
    "ASCENSORES": ("OTROS_GASTOS",),
    "CONTADORES": ("ACS", "CALEFACCION"),
}

@dataclass(frozen=True)
class EligibilityDecision:
    status: str
    reason: str | None
    concept_keys: tuple[str, ...] = ()
```

La función pura valida primero tipo, comunidad y período, y después exige intersección entre `SERVICE_MODULES[service_family]` y `profile.active_modules`.

- [ ] **Step 4: Persistir proveedor, evidencia y decisión**

En `_persist_analysis`, actualizar `provider_key`, `analysis_version`, `eligibility_status`, `eligibility_reason` y hacer `UPSERT` de cada `FieldEvidence`. Si la raíz es `provider_unknown`, crear solo `document.provider`; no crear simultáneamente incidencias de fecha e importe hasta reanalizar después de resolver el proveedor.

- [ ] **Step 5: Bloquear autoaplicación no elegible**

Antes de `confirm_source_candidates` en `_auto_apply_clean_analysis`:

```python
eligibility = connection.execute(
    "SELECT eligibility_status FROM source_documents WHERE id_document=?",
    (document.id_document,),
).fetchone()[0]
if eligibility != "eligible":
    return document
```

Para `not_applicable`, marcar `status='not_applicable'` y registrar el motivo en `case_history_events`; nunca insertar en `facturas`.

- [ ] **Step 6: Añadir regresión de persistencia y agrupación**

```python
def test_unknown_provider_creates_one_root_issue_not_three_missing_fields(self):
    result = case_ingestion.add_analysed_document_to_case(
        self.connection, self.case.id_case, source_path=self.source_path,
        archive_root=self.archive_root,
        analysis=SourceAnalysis("invoice", "medium", {},
            ("fecha_inicio", "fecha_fin", "importe_total"),
            review_message="Proveedor desconocido"),
    )
    issues = document_review.list_open_issues(self.connection, self.case.id_case)
    self.assertEqual(["document.provider"], [issue.field_name for issue in issues])
    self.assertEqual("under_review", result.document.status)
```

- [ ] **Step 7: Ejecutar pruebas de flujo**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_eligibility tests.test_expedient_flow tests.test_unified_ingestion_regressions -v
```

Expected: PASS.

- [ ] **Step 8: Commit**

```powershell
git add core/source_eligibility.py core/source_analysis.py core/case_ingestion.py core/document_review.py tests/test_source_eligibility.py tests/test_expedient_flow.py tests/test_unified_ingestion_regressions.py
git commit -m "feat: aplicar solo fuentes elegibles al expediente"
```

---

### Task 7: Ampliación auditada del catálogo y verificador privado

**Files:**
- Create: `scripts/audit_provider_catalog.py`
- Create: `tests/test_provider_catalog_audit.py`
- Modify: `config/proveedores.json`
- Modify: `.gitignore`
- Modify: `tests/test_provider_registry.py`

**Interfaces:**
- Consumes: carpetas locales, caché de texto y registro global.
- Produces: `audit_documents(paths, registry, text_loader) -> AuditReport`; métricas agregadas sin texto personal; perfiles para los proveedores recurrentes; salida privada en `.private-audit/provider-catalog.json`.

- [ ] **Step 1: Escribir pruebas del informe agregado**

```python
class ProviderCatalogAuditTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.directory = Path(self.temp.name)
        self.registry = provider_registry_from_payload({"proveedores": {
            key: {
                "tax_ids": [], "aliases": [key], "document_types": ["invoice"],
                "service_family": "MANTENIMIENTO",
                "extractor_family": "standard_spanish_invoice",
                "required_signatures": [r"\bFACTURA\b"],
                "excluded_signatures": [r"\bPRESUPUESTO\b"],
            }
            for key in ("TIERSAN", "ECHEMAN")
        }})

    def tearDown(self):
        self.temp.cleanup()

    def test_audit_report_contains_counts_but_no_document_text(self):
        path = self.directory / "tiersan.pdf"
        path.write_bytes(b"synthetic-pdf")
        report = audit_documents(
            (path,), registry=self.registry,
            text_loader=lambda _path: "TIERSAN FACTURA F-1 Total 121,00 €",
        )
        self.assertEqual(1, report.unique_documents)
        self.assertEqual(1, report.recognised_invoices)
        self.assertNotIn("TIERSAN FACTURA", json.dumps(asdict(report)))

    def test_duplicate_sha_is_counted_once(self):
        first = self.directory / "a.pdf"
        second = self.directory / "b.pdf"
        first.write_bytes(b"same-synthetic-pdf")
        second.write_bytes(b"same-synthetic-pdf")
        report = audit_documents(
            (first, second), registry=self.registry,
            text_loader=lambda _path: "ECHEMAN FACTURA F-1 Total 20,00 €",
        )
        self.assertEqual(1, report.unique_documents)
```

- [ ] **Step 2: Ejecutar y comprobar fallo**

Run: `.\.venv-fase1\Scripts\python.exe -m unittest tests.test_provider_catalog_audit -v`

Expected: ERROR porque el script aún no existe.

- [ ] **Step 3: Implementar auditor por SHA y métricas**

El informe debe contener solo: número de rutas, SHA únicos, tipo, proveedor clave, confianza, códigos de diagnóstico, duración y conteos agregados. Añadir `.private-audit/` a `.gitignore`.

CLI:

```powershell
.\.venv-fase1\Scripts\python.exe scripts\audit_provider_catalog.py --root "C:\ruta\de\fuentes" --database data\gestion.db --output .private-audit\provider-catalog.json
```

- [ ] **Step 4: Añadir los perfiles recurrentes por prioridad**

Incorporar claves globales independientes de comunidad para:

```text
TIERSAN, GOMEZ_GROUP_METERING (formatos adicionales), ECHEMAN,
FENIE_ENERGIA, CERRAJERA_MONCASI, LIMPIEZAS_COTE, SCHINDLER,
JPG_REPARACIONES_ELECTRICAS, LIMPIEZAS_UTEBO, ORONA, MOEVE,
TRITERMIA, TELESER, ISS, LABOIL, ISTA, ASCENSORS_SALES,
GAS_INSTALACIONES_MANTENIMIENTOS, VILAHEXDOSS, MARTINEZ_OTERO,
LIMPIEZAS_MOREDA, JARDINERIA_JESUS_GRACIA, LIMPIEZAS_MARCEN
```

Cada entrada declara alias exactos observados, tipos documentales permitidos, familia de servicio, familia extractora, firmas requeridas y exclusiones. El CIF/NIF se añade únicamente cuando el auditor lo extrae de forma inequívoca del bloque del emisor; un identificador dudoso no se guarda.

- [ ] **Step 5: Añadir pruebas sintéticas de colisión por cada familia**

Usar `subTest(provider_key=...)` y verificar como mínimo: alias + estructura de factura reconoce el proveedor; el mismo alias dentro de un presupuesto no produce factura; banco/cliente no reemplaza al emisor; dos coincidencias iguales devuelven ambigüedad.

- [ ] **Step 6: Ejecutar auditoría real local y comprobar umbrales**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe scripts\audit_provider_catalog.py --root "C:\Users\Jose\Proyectos\Soporte" --database data\gestion.db --output .private-audit\provider-catalog.json
```

Expected en el resumen de consola:

```text
unique_pdf=661
text_invoices_recognised_pct>=90.0
provider_decisions_remaining<60
non_invoice_false_positives=0
```

El número `unique_pdf` puede aumentar si el usuario añadió nuevos PDF; los tres umbrales de calidad no cambian.

- [ ] **Step 7: Ejecutar pruebas versionadas**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest tests.test_provider_registry tests.test_provider_catalog_audit tests.test_source_analysis -v
```

Expected: PASS.

- [ ] **Step 8: Commit**

```powershell
git add .gitignore config/proveedores.json core/provider_registry.py scripts/audit_provider_catalog.py tests/test_provider_registry.py tests/test_provider_catalog_audit.py tests/test_source_analysis.py
git commit -m "feat: ampliar catalogo global con auditoria privada"
```

---

### Task 8: Lotes paralelos, progreso por fases e interfaz de revisión

**Files:**
- Create: `core/source_batch.py`
- Create: `tests/test_source_batch.py`
- Modify: `core/expedient_ui.py:1655-1850,1852-2002`
- Modify: `core/community_onboarding.py:778-900`
- Modify: `tests/test_expedient_ui.py`
- Modify: `tests/test_community_onboarding.py`

**Interfaces:**
- Consumes: `analyse_source`, conexión por worker, rutas deduplicadas.
- Produces: `BatchProgress`; `analyse_batch(paths, *, community_code, database_path, max_workers=3, progress=None) -> tuple[BatchItem, ...]`.

- [ ] **Step 1: Escribir pruebas de continuidad, límite y progreso**

```python
def test_batch_continues_after_one_document_error(self):
    events = []
    def analyser(path, **_kwargs):
        if path.name == "broken.pdf":
            raise ValueError("PDF ilegible")
        return SourceAnalysis.invoice({
            "tipo_suministro": "GAS", "fecha_inicio": "2026-01-01",
            "fecha_fin": "2026-01-31", "importe_total": "10.00",
        })
    results = analyse_batch(
        (Path("ok.pdf"), Path("broken.pdf"), Path("ok2.pdf")),
        community_code="658", database_path=self.database_path,
        analyser=analyser, max_workers=2, progress=events.append,
    )
    self.assertEqual(3, len(results))
    self.assertEqual(1, sum(item.error is not None for item in results))
    self.assertEqual("completed", events[-1].phase)

def test_batch_never_uses_more_than_three_workers(self):
    with mock.patch("source_batch.ThreadPoolExecutor") as executor:
        analyse_batch((), community_code="658", database_path=self.database_path)
    executor.assert_called_once_with(max_workers=3, thread_name_prefix="source-analysis")
```

- [ ] **Step 2: Ejecutar y comprobar fallo**

Run: `.\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_batch -v`

Expected: ERROR porque el módulo aún no existe.

- [ ] **Step 3: Implementar lote deduplicado y seguro**

```python
@dataclass(frozen=True)
class BatchProgress:
    completed: int
    total: int
    filename: str
    phase: str  # hashing, text, classification, provider, fields, completed


@dataclass(frozen=True)
class BatchItem:
    path: Path
    analysis: SourceAnalysis | None
    error: str | None
```

Calcular SHA antes de crear workers, analizar una sola vez cada contenido duplicado y abrir una conexión con `busy_timeout=5000` dentro de cada worker. Ordenar el resultado final según las rutas de entrada para que la UI sea determinista.

- [ ] **Step 4: Integrar el lote en alta inicial y “Añadir fuentes”**

Sustituir el bucle de análisis secuencial por `analyse_batch`. Mantener las escrituras canónicas serializadas en el hilo de coordinación después de recibir cada `BatchItem`; la UI solo se actualiza mediante `app.after(0, ...)`.

- [ ] **Step 5: Mostrar fase, caché y decisión raíz**

La barra de estado usa textos concretos: `Huella`, `Leyendo texto`, `Clasificando`, `Identificando proveedor`, `Extrayendo campos`, `Finalizado`. La revisión muestra una tarjeta por documento; una decisión `document.provider` explica qué dato fiscal o nombre del emisor debe buscarse y permite abrir la página/localizador guardado.

- [ ] **Step 6: Ejecutar pruebas UI y lote**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest tests.test_source_batch tests.test_expedient_ui tests.test_community_onboarding -v
```

Expected: PASS.

- [ ] **Step 7: Commit**

```powershell
git add core/source_batch.py core/expedient_ui.py core/community_onboarding.py tests/test_source_batch.py tests/test_expedient_ui.py tests/test_community_onboarding.py
git commit -m "feat: analizar fuentes en lote con progreso"
```

---

### Task 9: Verificación integral y manual operativo

**Files:**
- Modify: `docs/manual-operacion.md`
- Modify only if a regression requires it: files owned by Tasks 1-8

**Interfaces:**
- Consumes: canalización completa y corpus local.
- Produces: instrucciones reproducibles, evidencia de umbrales y suite verde.

- [ ] **Step 1: Documentar el flujo visible para el usuario**

Añadir al manual:

```text
1. Añadir archivos o una carpeta completa.
2. El sistema deduplica, reutiliza texto/OCR y separa documentos por tipo.
3. Las fuentes seguras se aplican sin confirmación individual.
4. “Validar” muestra solo decisiones raíz pendientes.
5. Una factura no aplicable se conserva en el histórico, pero no altera el reparto.
6. Después de corregir proveedor o tipo, “Reanalizar fuentes” recalcula únicamente lo pendiente.
```

Incluir cómo ejecutar el auditor privado, dónde queda su JSON ignorado y cómo interpretar los tres umbrales.

- [ ] **Step 2: Ejecutar la suite completa en entorno limpio**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -v
```

Expected: todas las pruebas terminan en `OK`.

- [ ] **Step 3: Ejecutar controles estructurales**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe -m compileall -q core scripts tests
git diff --check
git status --short
```

Expected: `compileall` y `git diff --check` devuelven código `0`; `data/gestion.db` continúa fuera del staged area.

- [ ] **Step 4: Ejecutar auditoría privada final**

Run:

```powershell
.\.venv-fase1\Scripts\python.exe scripts\audit_provider_catalog.py --root "C:\Users\Jose\Proyectos\Soporte" --database data\gestion.db --output .private-audit\provider-catalog.json
```

Expected: reconocimiento ≥90 %, menos de 60 decisiones de proveedor, cero falsos positivos no-factura y ningún bloqueo del lote.

- [ ] **Step 5: Commit de documentación o correcciones finales**

```powershell
git add docs/manual-operacion.md
git commit -m "docs: explicar deteccion global de fuentes"
```

- [ ] **Step 6: Revisar el historial de la entrega**

Run:

```powershell
git log --oneline --max-count=12
git status --short
```

Expected: commits separados por responsabilidad y ningún archivo privado o `data/gestion.db` incluido.
