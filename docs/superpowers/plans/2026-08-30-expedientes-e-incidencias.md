# Expedientes e incidencias de regularización Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Permitir crear un expediente de una comunidad para cualquier rango de fechas, incorporar sus archivos fuente y obligar a corregir en pantalla los datos inciertos antes de que entren en cálculo.

**Architecture:** Una migración añade expedientes, documentos, candidatos de extracción, incidencias y correcciones sin alterar las tablas históricas existentes. `expedient_service.py` controla el ciclo de vida y los archivos con huella; `document_review.py` convierte los datos extraídos en incidencias y las resuelve con auditoría. La interfaz solo coordina diálogos y muestra estados; toda transición y validación permanece en servicios testeables.

**Tech Stack:** Python 3.14, SQLite, `hashlib`, `shutil`, `pathlib`, CustomTkinter 5.2.2, `unittest`.

**Spec:** [`docs/superpowers/specs/2026-08-30-nucleo-configurable-design.md`](../specs/2026-08-30-nucleo-configurable-design.md)

## Global Constraints

- No crear lógica que dependa de los códigos `658` o `644`; las comunidades son datos.
- Mantener `periodos`, `facturas`, `lecturas_vecino` y el flujo histórico sin renombrar ni borrar columnas.
- Importes y fechas manuales se almacenan como texto normalizado; el cálculo monetario llegará en la entrega de componentes configurables.
- No se considera válido un dato manual sin documento, campo, valor, motivo, fecha y responsable local `usuario_local`.
- Una incidencia abierta bloquea la transición del expediente a `ready_for_calculation`.
- Los archivos quedan archivados en `data/expedientes/<id>/fuentes/` mediante copia; el original no se mueve ni se elimina.
- La duplicación se decide por SHA-256 dentro de un expediente; añadir el mismo archivo dos veces no crea otro documento ni otra incidencia.
- No incorporar facturas ni datos personales reales en pruebas ni en Git; usar ficheros sintéticos temporales.
- Ejecutar pruebas desde `Flujo calculo regularizaciones` con `PYTHONUTF8=1` para no fallar al imprimir los mensajes existentes.
- No modificar ni versionar `Nuevo Documento de texto.txt` ni `plantillas/CARTA_PLANTILLA_DATOS_VARIABLES.docx`.

---

## Estructura de archivos acordada

| Archivo | Responsabilidad |
| --- | --- |
| `core/db_migrations.py` | Migración 2, tablas e índices de expedientes y revisión. |
| `core/expedient_models.py` | Enumeraciones y dataclasses inmutables que cruzan servicios e interfaz. |
| `core/expedient_service.py` | Crear, consultar y cambiar el estado de expedientes; archivar y deduplicar documentos. |
| `core/document_review.py` | Validar candidatos, crear/consultar incidencias y confirmar correcciones auditables. |
| `core/case_ingestion.py` | Controlador sin interfaz que registra un archivo, sus candidatos y sus incidencias requeridas. |
| `core/expedient_ui.py` | Diálogos CTk de crear expediente, añadir archivos y resolver una incidencia. |
| `core/app.py` | Enlace del panel nuevo a los servicios y actualización de los indicadores de progreso. |
| `tests/test_expedient_service.py` | Casos de fecha, ciclo de vida, archivos y deduplicación. |
| `tests/test_document_review.py` | Casos de incidencias, correcciones y bloqueo de cálculo. |
| `tests/test_expedient_flow.py` | Caso integral sintético, sin iniciar Tk. |

## Interfaces públicas de la entrega

```python
# core/expedient_models.py
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Literal

CaseStatus = Literal[
    "draft", "gathering_sources", "under_review", "ready_for_calculation",
    "calculated", "reconciled", "deliveries_generated", "closed",
]
DocumentStatus = Literal["registered", "under_review", "validated", "not_applicable"]
IssueStatus = Literal["open", "resolved", "dismissed"]

@dataclass(frozen=True)
class RegularizationCase:
    id_case: int
    community_id: int
    name: str
    start_date: date
    end_date: date
    status: CaseStatus

@dataclass(frozen=True)
class SourceDocument:
    id_document: int
    id_case: int
    original_name: str
    archived_path: Path
    sha256: str
    document_kind: str
    status: DocumentStatus

@dataclass(frozen=True)
class ReviewIssue:
    id_issue: int
    id_case: int
    id_document: int
    archived_path: Path
    code: str
    field_name: str
    detected_value: str | None
    message: str
    status: IssueStatus

# core/expedient_service.py
def create_case(connection: sqlite3.Connection, community_id: int, *,
                name: str, start_date: date, end_date: date) -> RegularizationCase: ...
def get_case(connection: sqlite3.Connection, case_id: int) -> RegularizationCase: ...
def list_cases(connection: sqlite3.Connection, community_id: int) -> tuple[RegularizationCase, ...]: ...
def register_source_document(connection: sqlite3.Connection, case_id: int, *,
                             source_path: str | Path, archive_root: str | Path,
                             document_kind: str) -> tuple[SourceDocument, bool]: ...
def set_case_status(connection: sqlite3.Connection, case_id: int,
                    status: CaseStatus) -> RegularizationCase: ...

# core/document_review.py
def record_candidates(connection: sqlite3.Connection, document_id: int,
                      candidates: Mapping[str, str | None], *, source: str) -> None: ...
def create_missing_field_issues(connection: sqlite3.Connection, case_id: int,
                                document_id: int, required_fields: Collection[str]) -> tuple[ReviewIssue, ...]: ...
def list_open_issues(connection: sqlite3.Connection, case_id: int) -> tuple[ReviewIssue, ...]: ...
def resolve_issue(connection: sqlite3.Connection, issue_id: int, *, value: str,
                  reason: str, resolved_by: str = "usuario_local") -> ReviewIssue: ...
def validate_case_ready(connection: sqlite3.Connection, case_id: int) -> RegularizationCase: ...
```

### Task 1: Añadir esquema de expedientes y auditoría de revisión

**Files:**
- Modify: `core/db_migrations.py`
- Test: `tests/test_db_migrations.py`

**Consumes:** `migrate(connection)` y las tablas base `comunidades` existentes.

**Produces:** versión de esquema 2, con las tablas `regularization_cases`, `source_documents`, `extraction_candidates`, `review_issues` y `manual_corrections`.

- [ ] **Step 1: Escribir las pruebas que describen la migración 2**

  Añadir a `tests/test_db_migrations.py` una prueba que cree las tablas base, ejecute `migrate(connection)` y exija las cinco tablas, sus columnas de integridad y la fila de versión 2:

  ```python
  def test_migration_two_creates_case_and_review_tables(self):
      with temporary_database() as (connection, _):
          for statement in gestor_bd.TABLAS:
              connection.execute(statement)
          connection.commit()
          version = gestor_bd.aplicar_migraciones(connection)
          tables = {row[0] for row in connection.execute(
              "SELECT name FROM sqlite_master WHERE type='table'"
          )}
          self.assertEqual(version, 2)
          self.assertTrue({
              "regularization_cases", "source_documents", "extraction_candidates",
              "review_issues", "manual_corrections",
          }.issubset(tables))
          columns = {row[1] for row in connection.execute(
              "PRAGMA table_info(review_issues)"
          )}
          self.assertTrue({"id_case", "id_document", "field_name", "status"}.issubset(columns))
  ```

  Añadir otra prueba que ejecute `migrate(connection)` dos veces y compruebe que `schema_migrations` contiene exactamente las versiones `1` y `2`.

- [ ] **Step 2: Ejecutar la prueba para observar el fallo esperado**

  Run:

  ```powershell
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest tests.test_db_migrations.DatabaseMigrationTest.test_migration_two_creates_case_and_review_tables -v
  ```

  Expected: fallo porque la versión actual es 1 y no existe `regularization_cases`.

- [ ] **Step 3: Implementar la migración mínima y transaccional**

  Elevar `CURRENT_SCHEMA_VERSION` a `2`, añadir `_migration_2(connection)` y registrarla en `MIGRATIONS`. Ejecutar estas sentencias, en este orden:

  ```sql
  CREATE TABLE regularization_cases (
      id_case INTEGER PRIMARY KEY AUTOINCREMENT,
      id_comunidad INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
      nombre TEXT NOT NULL,
      fecha_inicio TEXT NOT NULL,
      fecha_fin TEXT NOT NULL,
      estado TEXT NOT NULL CHECK(estado IN (
          'draft','gathering_sources','under_review','ready_for_calculation',
          'calculated','reconciled','deliveries_generated','closed'
      )),
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      CHECK(fecha_fin >= fecha_inicio),
      UNIQUE(id_comunidad, nombre)
  );
  CREATE TABLE source_documents (
      id_document INTEGER PRIMARY KEY AUTOINCREMENT,
      id_case INTEGER NOT NULL REFERENCES regularization_cases(id_case) ON DELETE CASCADE,
      original_name TEXT NOT NULL,
      archived_path TEXT NOT NULL,
      sha256 TEXT NOT NULL,
      document_kind TEXT NOT NULL,
      status TEXT NOT NULL CHECK(status IN ('registered','under_review','validated','not_applicable')),
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(id_case, sha256)
  );
  CREATE TABLE extraction_candidates (
      id_candidate INTEGER PRIMARY KEY AUTOINCREMENT,
      id_document INTEGER NOT NULL REFERENCES source_documents(id_document) ON DELETE CASCADE,
      field_name TEXT NOT NULL,
      value TEXT,
      source TEXT NOT NULL,
      validation_status TEXT NOT NULL CHECK(validation_status IN ('candidate','validated','rejected')),
      UNIQUE(id_document, field_name)
  );
  CREATE TABLE review_issues (
      id_issue INTEGER PRIMARY KEY AUTOINCREMENT,
      id_case INTEGER NOT NULL REFERENCES regularization_cases(id_case) ON DELETE CASCADE,
      id_document INTEGER NOT NULL REFERENCES source_documents(id_document) ON DELETE CASCADE,
      code TEXT NOT NULL,
      field_name TEXT NOT NULL,
      detected_value TEXT,
      message TEXT NOT NULL,
      status TEXT NOT NULL CHECK(status IN ('open','resolved','dismissed')),
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      resolved_at TEXT,
      UNIQUE(id_document, code, field_name, status)
  );
  CREATE TABLE manual_corrections (
      id_correction INTEGER PRIMARY KEY AUTOINCREMENT,
      id_issue INTEGER NOT NULL REFERENCES review_issues(id_issue) ON DELETE CASCADE,
      original_value TEXT,
      corrected_value TEXT NOT NULL,
      reason TEXT NOT NULL,
      resolved_by TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
  );
  CREATE INDEX idx_cases_community ON regularization_cases(id_comunidad, fecha_inicio, fecha_fin);
  CREATE INDEX idx_documents_case ON source_documents(id_case, status);
  CREATE INDEX idx_issues_case_open ON review_issues(id_case, status);
  ```

  Mantener el `BEGIN IMMEDIATE` y rollback que ya protege `migrate()`; no añadir commits dentro de `_migration_2`.

- [ ] **Step 4: Ejecutar la prueba y después todas las migraciones**

  Run:

  ```powershell
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest tests.test_db_migrations -v
  ```

  Expected: todas las pruebas de migraciones pasan.

- [ ] **Step 5: Confirmar el cambio en un commit aislado**

  ```powershell
  git add core/db_migrations.py tests/test_db_migrations.py
  git commit -m "Añade expedientes e incidencias auditables"
  ```

### Task 2: Crear expedientes de fechas libres y archivar fuentes

**Files:**
- Create: `core/expedient_models.py`
- Create: `core/expedient_service.py`
- Create: `tests/test_expedient_service.py`

**Consumes:** las tablas de la migración 2 y conexiones SQLite con `row_factory=sqlite3.Row`.

**Produces:** creación/listado de expedientes, transición controlada y documentos copiados/deduplicados por huella.

- [ ] **Step 1: Escribir pruebas de fechas, estados y deduplicación**

  Crear `tests/test_expedient_service.py`. Incluir una prueba de fechas libres que cree una comunidad sintética, cree el expediente `Invierno parcial` entre `2026-01-15` y `2026-03-14`, y exija que ambas fechas y `draft` se preservan. Incluir una prueba de orden inválido:

  ```python
  def test_create_case_rejects_end_before_start(self):
      with self.assertRaisesRegex(ValueError, "fin posterior"):
          expedient_service.create_case(
              self.connection, self.community_id, name="incorrecto",
              start_date=date(2026, 5, 1), end_date=date(2026, 4, 30),
          )
  ```

  Crear un PDF sintético con `Path.write_bytes(b"%PDF-1.4 prueba")`; registrarlo dos veces con el mismo `archive_root`. Exigir que la primera llamada devuelve `created=True`, la segunda `created=False`, ambos documentos tienen el mismo id y existe solo una copia bajo `<archive_root>/<id_case>/fuentes/`. Después del primer registro exigir que el expediente pasó de `draft` a `gathering_sources`.

  Añadir una prueba que intenta pasar directamente de `draft` a `ready_for_calculation` y exige `ValueError` con `gathering_sources` en el mensaje.

- [ ] **Step 2: Ejecutar la prueba para observar el fallo esperado**

  Run:

  ```powershell
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_service -v
  ```

  Expected: error de importación porque `expedient_service` todavía no existe.

- [ ] **Step 3: Implementar modelos y servicio mínimo**

  En `core/expedient_models.py`, definir exactamente los aliases y dataclasses del bloque «Interfaces públicas», usando `date.fromisoformat()` y `Path` en las conversiones de filas.

  En `core/expedient_service.py`:

  - Definir `ALLOWED_TRANSITIONS`:

    ```python
    ALLOWED_TRANSITIONS = {
        "draft": {"gathering_sources"},
        "gathering_sources": {"under_review"},
        "under_review": {"gathering_sources", "ready_for_calculation"},
        "ready_for_calculation": {"under_review", "calculated"},
        "calculated": {"reconciled", "under_review"},
        "reconciled": {"deliveries_generated", "under_review"},
        "deliveries_generated": {"closed", "under_review"},
        "closed": {"under_review"},
    }
    ```

  - Implementar `create_case()` validando `end_date >= start_date`, insertando `draft` y devolviendo el registro.
  - Implementar `get_case()` que lance `LookupError("Expediente no encontrado")` cuando no existe y `list_cases()` ordenado por `fecha_inicio DESC, id_case DESC`.
  - Implementar `set_case_status()` contra `ALLOWED_TRANSITIONS`; actualizar `updated_at` y devolver el estado nuevo.
  - Implementar `_sha256(path)` en bloques de 64 KiB y `register_source_document()`: validar archivo existente, encontrar el mismo hash del expediente antes de copiar, crear `<archive_root>/<case_id>/fuentes`, elegir `<sha256[:12]>_<nombre_saneado>`, copiar con `shutil.copy2`, insertar con estado `registered`, pasar `draft` a `gathering_sources` y devolver `(document, True)`.
  - Si la copia falla, borrar la copia parcial si existe, propagar la excepción y no insertar una fila.

- [ ] **Step 4: Ejecutar las pruebas del servicio**

  Run:

  ```powershell
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_service -v
  ```

  Expected: todas pasan; el segundo registro no crea duplicados y la transición inválida sigue bloqueada.

- [ ] **Step 5: Confirmar el cambio en un commit aislado**

  ```powershell
  git add core/expedient_models.py core/expedient_service.py tests/test_expedient_service.py
  git commit -m "Crea expedientes con fuentes trazables"
  ```

### Task 3: Convertir candidatos faltantes en incidencias y registrar correcciones

**Files:**
- Modify: `core/expedient_models.py`
- Create: `core/document_review.py`
- Create: `tests/test_document_review.py`

**Consumes:** `SourceDocument`, `ReviewIssue`, las tablas de migración 2 y los documentos registrados por `expedient_service`.

**Produces:** cola de incidencias, correcciones manuales persistentes y transición a `ready_for_calculation` bloqueada o permitida según la cola.

- [ ] **Step 1: Escribir pruebas de revisión humana obligatoria**

  Crear `tests/test_document_review.py` con un `setUp()` que migra una BD temporal, crea comunidad, expediente y documento sintético. Añadir estas pruebas:

  ```python
  def test_missing_required_candidate_creates_open_issue(self):
      document_review.record_candidates(
          self.connection, self.document.id_document,
          {"importe_total": "123.45", "fecha_inicio": None}, source="pdf",
      )
      issues = document_review.create_missing_field_issues(
          self.connection, self.case.id_case, self.document.id_document,
          required_fields=("importe_total", "fecha_inicio"),
      )
      self.assertEqual([(item.field_name, item.status) for item in issues],
                       [("fecha_inicio", "open")])
  ```

  Escribir otra prueba que resuelva la incidencia con `value="2026-01-01"` y `reason="Confirmado en la primera página"`, exija una fila de `manual_corrections`, un candidato `validated` para `fecha_inicio`, incidencia `resolved` y fecha de resolución no nula.

  Añadir otra prueba que llame `validate_case_ready()` con una incidencia abierta y espere `ValueError("1 incidencia abierta")`; tras resolverla, esperar estado `ready_for_calculation`.

- [ ] **Step 2: Ejecutar la prueba para observar el fallo esperado**

  Run:

  ```powershell
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest tests.test_document_review -v
  ```

  Expected: error de importación porque `document_review` todavía no existe.

- [ ] **Step 3: Implementar la cola de revisión**

  En `core/document_review.py`:

  - `record_candidates()` hará `INSERT ... ON CONFLICT(id_document, field_name) DO UPDATE`, guardará `candidate` si el valor es vacío/`None` y `validated` si contiene texto no vacío.
  - `create_missing_field_issues()` consultará cada campo requerido, creará `MISSING_REQUIRED_FIELD` solo si no hay valor no vacío y devolverá solo las incidencias abiertas creadas o existentes; antes de insertar cambiará el documento a `under_review` y el expediente de `gathering_sources` a `under_review`.
  - `list_open_issues()` devolverá `ReviewIssue` ordenados por `id_issue` y unirá `source_documents` para devolver `archived_path`, además de comprobar que el vínculo pertenece al caso pedido.
  - `resolve_issue()` rechazará valores vacíos y razones vacías con `ValueError`; insertará la corrección, actualizará/creará el candidato con `source='manual'` y `validation_status='validated'`, y actualizará la incidencia a `resolved` con `resolved_at`.
  - `validate_case_ready()` contará incidencias `open`; si existen levantará `ValueError(f"{count} incidencia abierta(s) por resolver")`. Si no existen, usará `set_case_status()` en pasos válidos: `draft → gathering_sources → under_review → ready_for_calculation` o solo el último salto desde `under_review`. Marcará todos los documentos del caso como `validated` antes de devolver el expediente.

  Cada función que modifique varias tablas abrirá `BEGIN IMMEDIATE` solo si no existe ya una transacción; confirmará en éxito y revertirá en error.

- [ ] **Step 4: Ejecutar las pruebas de revisión y la suite actual**

  Run:

  ```powershell
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest tests.test_document_review -v
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -t . -q
  ```

  Expected: el módulo nuevo y las pruebas históricas pasan sin modificar la semántica de `facturas` ni `periodos`.

- [ ] **Step 5: Confirmar el cambio en un commit aislado**

  ```powershell
  git add core/document_review.py tests/test_document_review.py
  git commit -m "Bloquea calculos con incidencias sin revisar"
  ```

### Task 4: Orquestar la entrada de documentos y enlazar la interfaz de expediente

**Files:**
- Create: `core/case_ingestion.py`
- Create: `core/expedient_ui.py`
- Modify: `core/app.py:74-105, 170-280, 769-850`
- Test: `tests/test_expedient_flow.py`

**Consumes:** los servicios de las tareas 2 y 3, `ui_moderna.py`, `filedialog`, `messagebox` y el panel de actividad existente.

**Produces:** una operación única y testeable de entrada documental, y las acciones de usuario «Crear expediente», «Añadir fuentes», «Resolver incidencias» con una lista de incidencias y su archivo.

- [ ] **Step 1: Escribir el caso integral de la nueva operación de entrada, sin Tk**

  Crear `tests/test_expedient_flow.py` que invoque `case_ingestion.add_document_to_case()`: crear expediente de seis meses, añadir un archivo `factura.pdf` junto a un único candidato `importe_total`, exigir una incidencia de `fecha_inicio`, comprobar que no está listo, resolverla y comprobar que termina `ready_for_calculation` con documentos validados. La prueba debe comprobar que la corrección conserva la cadena original `None` y la cadena confirmada, sin iniciar `AppGestionFincas`.

  Añadir en la misma clase un segundo registro del mismo archivo después de resolver la incidencia. Exigir un único `source_documents`, una única incidencia `resolved`, una única corrección y expediente `ready_for_calculation` después de invocar `validate_case_ready()` por segunda vez.

  La interfaz de producción será:

  ```python
  @dataclass(frozen=True)
  class IngestionResult:
      document: SourceDocument
      created: bool
      open_issue_count: int

  def add_document_to_case(connection: sqlite3.Connection, case_id: int, *,
                           source_path: str | Path, archive_root: str | Path,
                           document_kind: str, candidates: Mapping[str, str | None],
                           required_fields: Collection[str]) -> IngestionResult: ...
  ```

- [ ] **Step 2: Ejecutar la prueba para observar el fallo esperado**

  Run:

  ```powershell
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_flow -v
  ```

  Expected: error de importación porque `case_ingestion` todavía no existe.

- [ ] **Step 3: Implementar el controlador de entrada y los controladores visuales separados**

  Crear `core/case_ingestion.py`. `add_document_to_case()` llamará a `register_source_document()`, después a `record_candidates()` y a `create_missing_field_issues()`. Cuando el documento ya exista por hash, no sobrescribirá candidatos confirmados, no recreará incidencias resueltas y devolverá el conteo actual de incidencias abiertas. La función devolverá `IngestionResult` sin contener widgets ni SQL de interfaz.

  Crear `core/expedient_ui.py` con estas funciones que reciben `app`, evitando mantener SQL o reglas de negocio en widgets:

  ```python
  def open_create_case_dialog(app: "AppGestionFincas") -> None: ...
  def open_add_sources_dialog(app: "AppGestionFincas", case_id: int) -> None: ...
  def open_issue_dialog(app: "AppGestionFincas", issue: ReviewIssue) -> None: ...
  ```

  `open_create_case_dialog` pedirá nombre, fecha inicio y fecha fin en formato `DD/MM/AAAA`; validará el formato antes de llamar `create_case()`. La fecha final anterior mostrará «La fecha de fin debe ser posterior o igual a la de inicio». Al aceptar, recargará la lista de expedientes y escribirá en el log nombre, fechas y duración.

  `open_add_sources_dialog` permitirá seleccionar varios `*.pdf`, `*.xlsx`, `*.xls`, `*.csv` y llamará `add_document_to_case()` en un hilo por cada archivo. Para cada archivo creará candidatos de metadatos mínimos: `nombre_archivo` y `tipo_documento`; para tipos factura requerirá `fecha_inicio`, `fecha_fin` e `importe_total`, por lo que inicialmente quedarán abiertos hasta extraerlos o confirmarlos. Mostrará conteos de añadidos y duplicados sin mover los originales.

  `open_issue_dialog` incluirá «Abrir archivo» mediante `os.startfile(issue.archived_path)`, etiqueta de campo, valor detectado, explicación, un campo de valor y otro de motivo. «Guardar corrección» llamará `resolve_issue()` y después `validate_case_ready()`; si siguen otras incidencias, cerrará el modal y actualizará la bandeja indicando el número restante. Cancelar no escribe en base de datos.

  Modificar `core/app.py` para importar opcionalmente `expedient_service`, `document_review` y `expedient_ui`, añadir `self.id_expediente = None`, sustituir el texto «EJERCICIO» por «EXPEDIENTE / PERÍODO» y añadir botones «Crear expediente», «Añadir fuentes» y «Resolver incidencias». La tarjeta de estado muestra fechas, número de documentos y `N incidencias por resolver`. Mantener los botones históricos en «Otras acciones», sin quitarlos.

- [ ] **Step 4: Ejecutar el caso integral y la suite completa**

  Run:

  ```powershell
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest tests.test_expedient_flow -v
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -t . -q
  ```

  Expected: la prueba integral llega a `ready_for_calculation`; la suite no genera errores.

- [ ] **Step 5: Comprobación manual de interfaz y commit**

  Lanzar la aplicación con `INICIAR_REGULARIZACION.bat`, seleccionar una comunidad, crear un expediente de seis meses y añadir un PDF de prueba sin importe. Confirmar que la bandeja muestra la incidencia, que «Abrir archivo» abre la copia archivada y que guardar un dato lo audita y actualiza el contador.

  ```powershell
  git add core/expedient_ui.py core/app.py tests/test_expedient_flow.py
  git commit -m "Añade revision manual de fuentes en la app"
  ```

### Task 5: Verificación de aceptación y guía de uso de la primera entrega

**Files:**
- Create: `docs/validacion-expedientes.md`
- Modify: `GUIA_PROCESAR_TODO.txt`
- Test: `tests/test_expedient_flow.py`

**Consumes:** todas las interfaces anteriores.

**Produces:** instrucciones operativas y evidencia reproducible de que ningún dato incierto salta a cálculo.

- [ ] **Step 1: Documentar el funcionamiento de usuario**

  Crear `docs/validacion-expedientes.md` con el procedimiento exacto:

  ```text
  1. Abrir la app y seleccionar comunidad.
  2. Pulsar Crear expediente y elegir fecha inicial/final.
  3. Pulsar Añadir fuentes; los originales permanecen intactos.
  4. Abrir Resolver incidencias; cada incidencia muestra su copia archivada.
  5. Confirmar valor y motivo; el sistema guarda la auditoría.
  6. Esperar el estado Listo para cálculo antes de regenerar Excel o cartas.
  ```

  Actualizar `GUIA_PROCESAR_TODO.txt` para distinguir el flujo histórico de la nueva primera fase y especificar que todavía no calcula automáticamente cuotas desde los documentos pendientes de configurar.

- [ ] **Step 2: Ejecutar verificación final**

  Run:

  ```powershell
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m unittest discover -s tests -t . -q
  $env:PYTHONUTF8='1'; .\.venv-fase1\Scripts\python.exe -m py_compile core\app.py core\expedient_models.py core\expedient_service.py core\document_review.py core\expedient_ui.py
  git diff --check
  ```

  Expected: todas las pruebas pasan, los módulos compilan y el diff no contiene errores de espacios.

- [ ] **Step 3: Confirmar documentación y código en un commit aislado**

  ```powershell
  git add docs/validacion-expedientes.md GUIA_PROCESAR_TODO.txt tests/test_expedient_flow.py
  git commit -m "Documenta flujo de expedientes revisables"
  ```

## Revisión de cobertura y coherencia

| Requisito del diseño | Tarea que lo cubre |
| --- | --- |
| Expediente con comunidad y fechas arbitrarias | 1 y 2 |
| Fuentes originales archivadas y deduplicadas | 2 |
| Incidencia obligatoria con campo preciso | 3 |
| Corrección manual con valor, motivo y responsable | 3 |
| Archivo abierto desde una pantalla de revisión | 4 |
| Progreso e impedimento de avance con incidencias | 4 |
| Reintentos sin duplicados y guía operativa | 5 |

Las reglas de componentes, prorrateo económico, extracción de proveedores, Excel definitivo, cartas portables y validación 644 se ejecutarán en las entregas B–E ya delimitadas por la especificación. No pertenecen a esta primera entrega para evitar mezclar migraciones de auditoría con cálculos económicos.
