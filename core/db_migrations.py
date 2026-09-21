"""Migraciones incrementales y transaccionales de la base de datos."""

from __future__ import annotations

import sqlite3
from collections.abc import Callable


CURRENT_SCHEMA_VERSION = 13


MIGRATION_1_SQL = (
    """
    CREATE TABLE import_batches (
        id_batch INTEGER PRIMARY KEY AUTOINCREMENT,
        id_comunidad INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
        id_periodo INTEGER REFERENCES periodos(id_periodo),
        source_kind TEXT NOT NULL,
        source_path TEXT NOT NULL,
        source_sha256 TEXT NOT NULL,
        status TEXT NOT NULL CHECK(status IN ('processing','validated','failed')),
        error_message TEXT,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        validated_at TEXT,
        UNIQUE(id_comunidad, source_sha256)
    )
    """,
    """
    CREATE TABLE source_values (
        id_source_value INTEGER PRIMARY KEY AUTOINCREMENT,
        id_batch INTEGER NOT NULL REFERENCES import_batches(id_batch) ON DELETE CASCADE,
        entity_type TEXT NOT NULL,
        entity_key TEXT NOT NULL,
        field_name TEXT NOT NULL,
        sheet_name TEXT,
        cell_address TEXT,
        raw_value TEXT,
        normalized_value TEXT,
        validation_status TEXT NOT NULL DEFAULT 'validated',
        UNIQUE(id_batch, entity_type, entity_key, field_name, sheet_name, cell_address)
    )
    """,
    """
    CREATE TABLE regularization_concepts (
        concept_key TEXT PRIMARY KEY,
        label TEXT NOT NULL,
        service TEXT,
        unit TEXT,
        display_order INTEGER NOT NULL,
        active INTEGER NOT NULL DEFAULT 1
    )
    """,
    """
    CREATE TABLE owner_concept_results (
        id_result INTEGER PRIMARY KEY AUTOINCREMENT,
        id_propietario INTEGER NOT NULL REFERENCES propietarios(id_propietario),
        id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
        concept_key TEXT NOT NULL REFERENCES regularization_concepts(concept_key),
        consumption REAL,
        consumption_unit TEXT,
        billed_cents INTEGER NOT NULL,
        actual_cents INTEGER NOT NULL,
        difference_cents INTEGER NOT NULL,
        id_batch INTEGER REFERENCES import_batches(id_batch),
        status TEXT NOT NULL DEFAULT 'calculated',
        UNIQUE(id_propietario, id_periodo, concept_key)
    )
    """,
    """
    CREATE TABLE reconciliations (
        id_reconciliation INTEGER PRIMARY KEY AUTOINCREMENT,
        id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
        concept_key TEXT NOT NULL REFERENCES regularization_concepts(concept_key),
        reference_billed_cents INTEGER NOT NULL,
        calculated_billed_cents INTEGER NOT NULL,
        reference_actual_cents INTEGER NOT NULL,
        calculated_actual_cents INTEGER NOT NULL,
        difference_cents INTEGER NOT NULL,
        status TEXT NOT NULL CHECK(status IN ('cuadrado','descuadrado')),
        checked_at TEXT NOT NULL DEFAULT (datetime('now')),
        UNIQUE(id_periodo, concept_key)
    )
    """,
    """
    CREATE TABLE community_letter_settings (
        id_comunidad INTEGER PRIMARY KEY REFERENCES comunidades(id_comunidad),
        selected_concepts_json TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
    """,
)


CONCEPTS = (
    ("acs_fixed", "Cuota fija de ACS", "ACS", "€/mes", 10),
    ("acs_variable", "Consumo de ACS", "ACS", "m³", 20),
    ("heating_fixed", "Cuota fija de calefacción", "CALEFACCION", "€/mes", 30),
    ("heating_variable", "Consumo de calefacción", "CALEFACCION", "kWh", 40),
    ("extraordinary_expense", "Gasto extraordinario", None, "€", 50),
    ("adjustment", "Ajuste", None, "€", 60),
    ("credit", "Abono", None, "€", 70),
    ("other", "Otros", None, "€", 80),
)


def _migration_1(connection: sqlite3.Connection) -> None:
    columns = {
        row[1] for row in connection.execute("PRAGMA table_info(propietarios)")
    }
    if "email" not in columns:
        connection.execute("ALTER TABLE propietarios ADD COLUMN email TEXT")

    for statement in MIGRATION_1_SQL:
        connection.execute(statement)

    connection.executemany(
        """
        INSERT INTO regularization_concepts
            (concept_key, label, service, unit, display_order)
        VALUES (?, ?, ?, ?, ?)
        """,
        CONCEPTS,
    )


def _migration_2(connection: sqlite3.Connection) -> None:
    statements = (
        """
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
        )
        """,
        """
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
        )
        """,
        """
        CREATE TABLE extraction_candidates (
            id_candidate INTEGER PRIMARY KEY AUTOINCREMENT,
            id_document INTEGER NOT NULL REFERENCES source_documents(id_document) ON DELETE CASCADE,
            field_name TEXT NOT NULL,
            value TEXT,
            source TEXT NOT NULL,
            validation_status TEXT NOT NULL CHECK(validation_status IN ('candidate','validated','rejected')),
            UNIQUE(id_document, field_name)
        )
        """,
        """
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
        )
        """,
        """
        CREATE TABLE manual_corrections (
            id_correction INTEGER PRIMARY KEY AUTOINCREMENT,
            id_issue INTEGER NOT NULL REFERENCES review_issues(id_issue) ON DELETE CASCADE,
            original_value TEXT,
            corrected_value TEXT NOT NULL,
            reason TEXT NOT NULL,
            resolved_by TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT (datetime('now'))
        )
        """,
        """
        CREATE INDEX idx_cases_community ON regularization_cases(id_comunidad, fecha_inicio, fecha_fin)
        """,
        """
        CREATE INDEX idx_documents_case ON source_documents(id_case, status)
        """,
        """
        CREATE INDEX idx_issues_case_open ON review_issues(id_case, status)
        """,
    )
    for statement in statements:
        connection.execute(statement)


def _migration_3(connection: sqlite3.Connection) -> None:
    """Vincula expedientes y periodos y añade la auditoría de salidas."""
    statements = (
        """
        ALTER TABLE regularization_cases
            ADD COLUMN id_periodo INTEGER REFERENCES periodos(id_periodo)
        """,
        """
        CREATE UNIQUE INDEX idx_case_period
            ON regularization_cases(id_comunidad, id_periodo)
            WHERE id_periodo IS NOT NULL
        """,
        """
        CREATE TABLE invoice_components (
            id_component INTEGER PRIMARY KEY AUTOINCREMENT,
            id_factura INTEGER NOT NULL REFERENCES facturas(id_factura) ON DELETE CASCADE,
            component_key TEXT NOT NULL,
            amount REAL NOT NULL,
            unit TEXT,
            source_sheet TEXT,
            source_cell TEXT,
            UNIQUE(id_factura, component_key)
        )
        """,
        """
        CREATE TABLE period_parameters (
            id_parameter INTEGER PRIMARY KEY AUTOINCREMENT,
            id_comunidad INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
            id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
            parameter_key TEXT NOT NULL,
            numeric_value REAL,
            text_value TEXT,
            unit TEXT,
            source_sheet TEXT,
            source_cell TEXT,
            UNIQUE(id_comunidad, id_periodo, parameter_key)
        )
        """,
        """
        CREATE TABLE excel_template_profiles (
            id_template_profile INTEGER PRIMARY KEY AUTOINCREMENT,
            id_comunidad INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
            profile_key TEXT NOT NULL,
            profile_version TEXT NOT NULL,
            template_relative_path TEXT NOT NULL,
            template_sha256 TEXT NOT NULL,
            profile_sha256 TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'active'
                CHECK(status IN ('active','inactive','retired')),
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            retired_at TEXT,
            UNIQUE(id_comunidad, profile_key, profile_version, template_sha256)
        )
        """,
        """
        CREATE TABLE excel_export_runs (
            id_export_run INTEGER PRIMARY KEY AUTOINCREMENT,
            id_case INTEGER NOT NULL REFERENCES regularization_cases(id_case),
            id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
            id_template_profile INTEGER NOT NULL
                REFERENCES excel_template_profiles(id_template_profile),
            input_sha256 TEXT NOT NULL,
            template_sha256 TEXT NOT NULL,
            output_path TEXT,
            diagnostic_path TEXT,
            status TEXT NOT NULL DEFAULT 'pending' CHECK(status IN (
                'pending','validating','generating','recalculating',
                'validated','published','failed'
            )),
            error_message TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            updated_at TEXT NOT NULL DEFAULT (datetime('now')),
            published_at TEXT
        )
        """,
        """
        CREATE TABLE letter_generation_runs (
            id_letter_run INTEGER PRIMARY KEY AUTOINCREMENT,
            id_case INTEGER NOT NULL REFERENCES regularization_cases(id_case),
            id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
            id_export_run INTEGER NOT NULL REFERENCES excel_export_runs(id_export_run),
            input_sha256 TEXT NOT NULL,
            template_sha256 TEXT NOT NULL,
            output_path TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'pending' CHECK(status IN (
                'pending','running','generating','completed','incomplete','failed'
            )),
            error_message TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            updated_at TEXT NOT NULL DEFAULT (datetime('now')),
            completed_at TEXT
        )
        """,
        """
        CREATE TABLE generated_letters (
            id_generated_letter INTEGER PRIMARY KEY AUTOINCREMENT,
            id_letter_run INTEGER NOT NULL
                REFERENCES letter_generation_runs(id_letter_run) ON DELETE CASCADE,
            id_propietario INTEGER NOT NULL REFERENCES propietarios(id_propietario),
            input_sha256 TEXT NOT NULL,
            template_sha256 TEXT NOT NULL,
            output_path TEXT,
            status TEXT NOT NULL DEFAULT 'pending'
                CHECK(status IN ('pending','generated','failed')),
            error_message TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            updated_at TEXT NOT NULL DEFAULT (datetime('now')),
            UNIQUE(id_letter_run, id_propietario)
        )
        """,
        """
        CREATE INDEX idx_export_runs_case_status
            ON excel_export_runs(id_case, status, created_at)
        """,
        """
        CREATE INDEX idx_letter_runs_case_status
            ON letter_generation_runs(id_case, status, created_at)
        """,
    )
    for statement in statements:
        connection.execute(statement)


def _migration_4(connection: sqlite3.Connection) -> None:
    """Enlaza el lote físico con cada expediente y uso lógico de la fuente."""
    statements = (
        """
        CREATE TABLE case_import_batches (
            id_case INTEGER NOT NULL
                REFERENCES regularization_cases(id_case) ON DELETE CASCADE,
            id_batch INTEGER NOT NULL
                REFERENCES import_batches(id_batch) ON DELETE CASCADE,
            id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
            source_kind TEXT NOT NULL,
            source_sha256 TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT (datetime('now')),
            PRIMARY KEY(id_case, id_batch, id_periodo, source_kind),
            UNIQUE(id_case, id_periodo, source_kind, source_sha256)
        )
        """,
        """
        CREATE INDEX idx_case_import_batches_period
            ON case_import_batches(id_case, id_periodo, source_kind)
        """,
    )
    for statement in statements:
        connection.execute(statement)


def _migration_5(connection: sqlite3.Connection) -> None:
    """Audita la aprobación explícita de lecturas estimadas.

    Una nota libre no basta para convertir una lectura irregular en consumo
    repartible: el nombre de quien la aprueba y el momento de aprobación son
    datos distintos y obligatorios para el motor de reparto de expedientes.
    """
    columns = {row[1] for row in connection.execute("PRAGMA table_info(lecturas_vecino)")}
    if "approved_by" not in columns:
        connection.execute("ALTER TABLE lecturas_vecino ADD COLUMN approved_by TEXT")
    if "approved_at" not in columns:
        connection.execute("ALTER TABLE lecturas_vecino ADD COLUMN approved_at TEXT")


def _migration_6(connection: sqlite3.Connection) -> None:
    """Conserva la trazabilidad de la clasificación y de cada extracción."""
    document_columns = {
        row[1] for row in connection.execute("PRAGMA table_info(source_documents)")
    }
    if "classification_confidence" not in document_columns:
        connection.execute(
            "ALTER TABLE source_documents ADD COLUMN classification_confidence TEXT"
        )

    candidate_columns = {
        row[1] for row in connection.execute("PRAGMA table_info(extraction_candidates)")
    }
    if "source_context" not in candidate_columns:
        connection.execute(
            "ALTER TABLE extraction_candidates ADD COLUMN source_context TEXT"
        )


def _migration_7(connection: sqlite3.Connection) -> None:
    """Distingue las incidencias regenerables de las creadas por una persona."""
    columns = {row[1] for row in connection.execute("PRAGMA table_info(review_issues)")}
    if "origin" not in columns:
        connection.execute(
            """ALTER TABLE review_issues
               ADD COLUMN origin TEXT NOT NULL DEFAULT 'manual'
               CHECK(origin IN ('automatic', 'manual'))"""
        )
        _reconcile_legacy_generated_issues(connection)


def _reconcile_legacy_generated_issues(connection: sqlite3.Connection) -> None:
    """Recognize the old generator's exact signature, never a reviewed outcome."""
    connection.execute(
        """UPDATE review_issues SET origin='automatic'
           WHERE origin='manual' AND status='open' AND resolved_at IS NULL
             AND code='MISSING_REQUIRED_FIELD' AND detected_value IS NULL
             AND field_name IN ('fecha_inicio','fecha_fin','importe_total')
             AND message='Falta el campo requerido: ' || field_name
             AND NOT EXISTS (SELECT 1 FROM schema_migrations m
                             WHERE m.version=7 AND review_issues.created_at>=m.applied_at)
             AND NOT EXISTS (SELECT 1 FROM manual_corrections c
                             WHERE c.id_issue=review_issues.id_issue)
             AND EXISTS (SELECT 1 FROM source_documents d
                         WHERE d.id_document=review_issues.id_document
                           AND d.classification_confidence IS NULL)
             AND NOT EXISTS (SELECT 1 FROM extraction_candidates c
                             WHERE c.id_document=review_issues.id_document
                               AND c.field_name=review_issues.field_name
                               AND (c.source LIKE 'manual%' OR c.validation_status='rejected'))"""
    )


def _migration_8(connection: sqlite3.Connection) -> None:
    _reconcile_legacy_generated_issues(connection)
    columns = {row[1] for row in connection.execute("PRAGMA table_info(source_documents)")}
    for name in ("source_context", "confirmed_by", "confirmed_at"):
        if name not in columns:
            connection.execute(f"ALTER TABLE source_documents ADD COLUMN {name} TEXT")
    connection.execute("""CREATE TABLE IF NOT EXISTS reading_periods (
        id_lectura INTEGER NOT NULL REFERENCES lecturas_vecino(id_lectura) ON DELETE CASCADE,
        id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
        PRIMARY KEY (id_lectura,id_periodo))""")
    connection.execute("""CREATE TABLE IF NOT EXISTS counter_reset_targets (
        id_issue INTEGER PRIMARY KEY REFERENCES review_issues(id_issue) ON DELETE CASCADE,
        initial_reading_id INTEGER NOT NULL REFERENCES lecturas_vecino(id_lectura),
        final_reading_id INTEGER NOT NULL REFERENCES lecturas_vecino(id_lectura))""")
    # The legacy period stays on the original row. Additional associations share
    # that same reading (including its explicit approval) without moving it.
    connection.execute("""CREATE VIEW IF NOT EXISTS period_readings AS
        SELECT id_lectura,id_propietario,id_periodo,tipo,fecha_lectura,
               valor_acumulado,estado,metodo_estimacion,fuente,notas,
               approved_by,approved_at FROM lecturas_vecino
        UNION
        SELECT l.id_lectura,l.id_propietario,p.id_periodo,l.tipo,l.fecha_lectura,
               l.valor_acumulado,l.estado,l.metodo_estimacion,l.fuente,l.notas,
               l.approved_by,l.approved_at
        FROM lecturas_vecino l JOIN reading_periods p ON p.id_lectura=l.id_lectura""")


def _migration_9(connection: sqlite3.Connection) -> None:
    """Keep an immutable, per-case timeline without duplicating source data."""
    connection.execute("""CREATE TABLE IF NOT EXISTS case_history_events (
        id_event INTEGER PRIMARY KEY AUTOINCREMENT,
        id_case INTEGER NOT NULL REFERENCES regularization_cases(id_case) ON DELETE CASCADE,
        id_document INTEGER REFERENCES source_documents(id_document) ON DELETE SET NULL,
        event_type TEXT NOT NULL,
        details_json TEXT NOT NULL DEFAULT '{}',
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )""")
    connection.execute("""CREATE INDEX IF NOT EXISTS idx_case_history_case_date
        ON case_history_events(id_case, created_at DESC)""")
    connection.execute("""CREATE TRIGGER IF NOT EXISTS history_source_registered
        AFTER INSERT ON source_documents
        BEGIN
            INSERT INTO case_history_events(id_case, id_document, event_type, details_json)
            VALUES (NEW.id_case, NEW.id_document, 'source_registered',
                json_object('name', NEW.original_name, 'kind', NEW.document_kind,
                            'sha256', NEW.sha256, 'status', NEW.status));
        END""")
    connection.execute("""CREATE TRIGGER IF NOT EXISTS history_source_status_changed
        AFTER UPDATE OF status ON source_documents
        WHEN OLD.status <> NEW.status
        BEGIN
            INSERT INTO case_history_events(id_case, id_document, event_type, details_json)
            VALUES (NEW.id_case, NEW.id_document, 'source_status_changed',
                json_object('from', OLD.status, 'to', NEW.status));
        END""")
    connection.execute("""CREATE TRIGGER IF NOT EXISTS history_manual_correction
        AFTER INSERT ON manual_corrections
        BEGIN
            INSERT INTO case_history_events(id_case, id_document, event_type, details_json)
            SELECT issue.id_case, issue.id_document, 'manual_correction',
                json_object('issue_id', issue.id_issue, 'field', issue.field_name,
                            'code', issue.code, 'from', NEW.original_value,
                            'to', NEW.corrected_value, 'reason', NEW.reason,
                            'by', NEW.resolved_by)
            FROM review_issues AS issue WHERE issue.id_issue = NEW.id_issue;
        END""")
    connection.execute("""CREATE TRIGGER IF NOT EXISTS history_case_status_changed
        AFTER UPDATE OF estado ON regularization_cases
        WHEN OLD.estado <> NEW.estado
        BEGIN
            INSERT INTO case_history_events(id_case, event_type, details_json)
            VALUES (NEW.id_case, 'case_status_changed',
                json_object('from', OLD.estado, 'to', NEW.estado));
        END""")
    connection.execute("""CREATE TRIGGER IF NOT EXISTS history_export_created
        AFTER INSERT ON excel_export_runs
        BEGIN
            INSERT INTO case_history_events(id_case, event_type, details_json)
            VALUES (NEW.id_case, 'excel_export_recorded',
                json_object('export_run_id', NEW.id_export_run, 'status', NEW.status,
                            'output_path', NEW.output_path));
        END""")
    connection.execute("""CREATE TRIGGER IF NOT EXISTS history_letter_run_created
        AFTER INSERT ON letter_generation_runs
        BEGIN
            INSERT INTO case_history_events(id_case, event_type, details_json)
            VALUES (NEW.id_case, 'letters_recorded',
                json_object('letter_run_id', NEW.id_letter_run, 'status', NEW.status,
                            'output_path', NEW.output_path));
        END""")
    # Databases created before this migration already retain their canonical
    # rows.  Seed their timeline once so historical periods are equally
    # reviewable; the NOT EXISTS guards make this safe on every later startup.
    connection.execute("""INSERT INTO case_history_events
            (id_case, id_document, event_type, details_json, created_at)
        SELECT d.id_case, d.id_document, 'source_registered',
               json_object('name', d.original_name, 'kind', d.document_kind,
                           'sha256', d.sha256, 'status', d.status), d.created_at
          FROM source_documents d
         WHERE NOT EXISTS (
             SELECT 1 FROM case_history_events h
              WHERE h.id_document=d.id_document AND h.event_type='source_registered'
         )""")
    connection.execute("""INSERT INTO case_history_events
            (id_case, id_document, event_type, details_json, created_at)
        SELECT i.id_case, i.id_document, 'manual_correction',
               json_object('issue_id', i.id_issue, 'field', i.field_name,
                           'code', i.code, 'from', c.original_value,
                           'to', c.corrected_value, 'reason', c.reason,
                           'by', c.resolved_by), c.created_at
          FROM manual_corrections c
          JOIN review_issues i ON i.id_issue=c.id_issue
         WHERE NOT EXISTS (
             SELECT 1 FROM case_history_events h
              WHERE h.event_type='manual_correction'
                AND json_extract(h.details_json, '$.issue_id')=i.id_issue
         )""")
    connection.execute("""INSERT INTO case_history_events
            (id_case, event_type, details_json, created_at)
        SELECT e.id_case, 'excel_export_recorded',
               json_object('export_run_id', e.id_export_run, 'status', e.status,
                           'output_path', e.output_path), e.created_at
          FROM excel_export_runs e
         WHERE NOT EXISTS (
             SELECT 1 FROM case_history_events h
              WHERE h.id_case=e.id_case AND h.event_type='excel_export_recorded'
                AND json_extract(h.details_json, '$.export_run_id')=e.id_export_run
         )""")
    connection.execute("""INSERT INTO case_history_events
            (id_case, event_type, details_json, created_at)
        SELECT l.id_case, 'letters_recorded',
               json_object('letter_run_id', l.id_letter_run, 'status', l.status,
                           'output_path', l.output_path), l.created_at
          FROM letter_generation_runs l
         WHERE NOT EXISTS (
             SELECT 1 FROM case_history_events h
              WHERE h.id_case=l.id_case AND h.event_type='letters_recorded'
                AND json_extract(h.details_json, '$.letter_run_id')=l.id_letter_run
         )""")


def _migration_10(connection: sqlite3.Connection) -> None:
    """Conserva el historial de incidencias y cada observación de contador."""
    connection.execute("""CREATE TABLE review_issues_rebuilt (
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
        origin TEXT NOT NULL DEFAULT 'manual' CHECK(origin IN ('automatic','manual'))
    )""")
    connection.execute("""INSERT INTO review_issues_rebuilt
        (id_issue,id_case,id_document,code,field_name,detected_value,message,status,
         created_at,resolved_at,origin)
        SELECT id_issue,id_case,id_document,code,field_name,detected_value,message,status,
               created_at,resolved_at,origin
          FROM review_issues""")
    connection.execute("DROP TRIGGER IF EXISTS history_manual_correction")
    connection.execute("DROP TABLE review_issues")
    connection.execute("ALTER TABLE review_issues_rebuilt RENAME TO review_issues")
    connection.execute("""CREATE INDEX idx_issues_case_open
        ON review_issues(id_case, status)""")
    connection.execute("""CREATE UNIQUE INDEX idx_review_issues_open_unique
        ON review_issues(id_document, code, field_name)
        WHERE status='open'""")
    connection.execute("""CREATE TRIGGER history_manual_correction
        AFTER INSERT ON manual_corrections
        BEGIN
            INSERT INTO case_history_events(id_case, id_document, event_type, details_json)
            SELECT issue.id_case, issue.id_document, 'manual_correction',
                json_object('issue_id', issue.id_issue, 'field', issue.field_name,
                            'code', issue.code, 'from', NEW.original_value,
                            'to', NEW.corrected_value, 'reason', NEW.reason,
                            'by', NEW.resolved_by)
            FROM review_issues AS issue WHERE issue.id_issue = NEW.id_issue;
        END""")
    connection.execute("""CREATE TABLE IF NOT EXISTS reading_observations (
        id_observation INTEGER PRIMARY KEY AUTOINCREMENT,
        id_propietario INTEGER NOT NULL REFERENCES propietarios(id_propietario),
        tipo TEXT NOT NULL,
        fecha_lectura TEXT NOT NULL,
        observed_value REAL NOT NULL,
        source_path TEXT NOT NULL,
        id_document INTEGER REFERENCES source_documents(id_document) ON DELETE SET NULL,
        status TEXT NOT NULL CHECK(status IN ('observed','carried_forward','conflict','review_required')),
        effective_reading_id INTEGER REFERENCES lecturas_vecino(id_lectura) ON DELETE SET NULL,
        previous_reading_id INTEGER REFERENCES lecturas_vecino(id_lectura) ON DELETE SET NULL,
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )""")
    connection.execute("""CREATE INDEX IF NOT EXISTS idx_reading_observations_owner_service_date
        ON reading_observations(id_propietario, tipo, fecha_lectura)""")


def _migration_11(connection: sqlite3.Connection) -> None:
    """Una fila por unidad, con su tipo, y sin duplicados por espacios.

    Los listados de propietarios llegaban de fuentes distintas y el mismo piso
    entraba dos veces ('MP-1ºA' y 'MP-1ºA        '), duplicando coeficientes y
    repartiendo las lecturas entre las dos copias. Aquí se fusionan las copias
    conservando la que tiene lecturas, se normaliza el código y se marca qué
    unidades son garaje o local, que no entran en la regularización.

    La normalización se escribe entera aquí, sin importar ayudas de otros
    módulos, para que la migración siga haciendo lo mismo el día que aquéllas
    cambien.
    """
    columnas = {
        row[1] for row in connection.execute("PRAGMA table_info(propietarios)")
    }
    if "tipo_unidad" not in columnas:
        connection.execute(
            "ALTER TABLE propietarios ADD COLUMN tipo_unidad TEXT NOT NULL DEFAULT 'vivienda'"
        )

    def normalizar(valor) -> str:
        return " ".join(str(valor or "").split())

    def clasificar(codigo: str) -> str:
        texto = codigo.upper()
        if any(marca in texto for marca in ("GAR", "PK", "PARKING", "APARCAMIENTO", "TRASTERO")):
            return "garaje"
        if any(marca in texto for marca in ("LOC", "COMERCIAL")):
            return "local"
        return "vivienda"

    hijas = (
        ("lecturas_vecino", "id_propietario"),
        ("repartos", "id_propietario"),
        ("owner_concept_results", "id_propietario"),
        ("generated_letters", "id_propietario"),
        ("reading_observations", "id_propietario"),
    )
    existentes = {
        row[0] for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        )
    }

    grupos: dict[tuple[int, str], list[int]] = {}
    for id_propietario, id_comunidad, codigo in connection.execute(
        "SELECT id_propietario,id_comunidad,codigo_vivienda FROM propietarios ORDER BY id_propietario"
    ):
        grupos.setdefault((id_comunidad, normalizar(codigo)), []).append(id_propietario)

    for (_comunidad, codigo), identificadores in grupos.items():
        if len(identificadores) > 1:
            def lecturas(identificador: int) -> int:
                if "lecturas_vecino" not in existentes:
                    return 0
                return connection.execute(
                    "SELECT COUNT(*) FROM lecturas_vecino WHERE id_propietario=?",
                    (identificador,),
                ).fetchone()[0]

            conservado = max(identificadores, key=lambda item: (lecturas(item), -item))
            for descartado in identificadores:
                if descartado == conservado:
                    continue
                for tabla, columna in hijas:
                    if tabla in existentes:
                        connection.execute(
                            f"UPDATE OR IGNORE {tabla} SET {columna}=? WHERE {columna}=?",
                            (conservado, descartado),
                        )
                connection.execute(
                    "DELETE FROM propietarios WHERE id_propietario=?", (descartado,)
                )
        else:
            conservado = identificadores[0]
        connection.execute(
            "UPDATE propietarios SET codigo_vivienda=?, tipo_unidad=? WHERE id_propietario=?",
            (codigo, clasificar(codigo), conservado),
        )


def _migration_12(connection: sqlite3.Connection) -> None:
    """Libro de lo cobrado a los propietarios por ACS y calefacción.

    Es el lado de los ingresos del estudio: las cuotas fijas mensuales que la
    comunidad gira y las liquidaciones variables de cada lectura. No sale de
    ninguna factura ni de ningún contador —lo decide la comunidad—, así que
    hasta ahora no había dónde guardarlo y la hoja de cobros quedaba vacía.

    De aquí salen el 'importe cobrado' del análisis (su fila 66) y, por
    diferencia con el coste real, la regularización de cada propietario.
    """
    connection.execute("""CREATE TABLE IF NOT EXISTS cuotas_servicio (
        id_cuota INTEGER PRIMARY KEY AUTOINCREMENT,
        id_comunidad INTEGER NOT NULL REFERENCES comunidades(id_comunidad),
        id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
        servicio TEXT NOT NULL CHECK(servicio IN ('ACS','CALEFACCION')),
        concepto TEXT NOT NULL CHECK(concepto IN ('fija','variable')),
        fecha TEXT NOT NULL,
        importe REAL NOT NULL,
        consumo REAL,
        fecha_inicio TEXT,
        fecha_fin TEXT,
        notas TEXT,
        creado_en TEXT NOT NULL DEFAULT (datetime('now')),
        UNIQUE(id_comunidad, id_periodo, servicio, concepto, fecha)
    )""")
    connection.execute("""CREATE INDEX IF NOT EXISTS idx_cuotas_periodo
        ON cuotas_servicio(id_comunidad, id_periodo, servicio, fecha)""")


def _migration_13(connection: sqlite3.Connection) -> None:
    """Conserva cada reparto y los coeficientes exactos que lo originaron.

    ``owner_concept_results`` continúa siendo la proyección vigente que usa la
    interfaz. Estas tablas son el historial inmutable necesario para poder
    reconstruir una regularización antigua aunque después cambie un
    propietario, un coeficiente o una regla del perfil.
    """
    connection.execute("""CREATE TABLE IF NOT EXISTS distribution_runs (
        id_distribution_run INTEGER PRIMARY KEY AUTOINCREMENT,
        id_case INTEGER NOT NULL REFERENCES regularization_cases(id_case),
        id_periodo INTEGER NOT NULL REFERENCES periodos(id_periodo),
        input_sha256 TEXT NOT NULL,
        status TEXT NOT NULL CHECK(status IN ('running','completed','failed')),
        error_message TEXT,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        completed_at TEXT
    )""")
    connection.execute("""CREATE INDEX IF NOT EXISTS idx_distribution_runs_case
        ON distribution_runs(id_case, created_at DESC, id_distribution_run DESC)""")
    connection.execute("""CREATE TABLE IF NOT EXISTS owner_distribution_snapshots (
        id_snapshot INTEGER PRIMARY KEY AUTOINCREMENT,
        id_distribution_run INTEGER NOT NULL
            REFERENCES distribution_runs(id_distribution_run) ON DELETE CASCADE,
        id_propietario INTEGER NOT NULL REFERENCES propietarios(id_propietario),
        concept_key TEXT NOT NULL REFERENCES regularization_concepts(concept_key),
        coefficient_raw REAL,
        coefficient_eligible_total REAL,
        coefficient_applied REAL,
        consumption REAL,
        consumption_unit TEXT,
        billed_cents INTEGER NOT NULL,
        actual_cents INTEGER NOT NULL,
        difference_cents INTEGER NOT NULL,
        UNIQUE(id_distribution_run, id_propietario, concept_key)
    )""")
    connection.execute("""CREATE INDEX IF NOT EXISTS idx_distribution_snapshots_owner
        ON owner_distribution_snapshots(id_propietario, id_distribution_run)""")


MIGRATIONS: dict[int, Callable[[sqlite3.Connection], None]] = {
    1: _migration_1,
    2: _migration_2,
    3: _migration_3,
    4: _migration_4,
    5: _migration_5,
    6: _migration_6,
    7: _migration_7,
    8: _migration_8,
    9: _migration_9,
    10: _migration_10,
    11: _migration_11,
    12: _migration_12,
    13: _migration_13,
}


def migrate(connection: sqlite3.Connection) -> int:
    """Aplica en orden las migraciones pendientes o revierte todo el lote."""
    if connection.in_transaction:
        connection.commit()

    applied = {
        row[0] for row in connection.execute("SELECT version FROM schema_migrations")
    } if connection.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='schema_migrations'"
    ).fetchone() else set()
    rebuild_review_issues = 10 not in applied
    foreign_keys_enabled = bool(connection.execute("PRAGMA foreign_keys").fetchone()[0])
    if rebuild_review_issues and foreign_keys_enabled:
        connection.execute("PRAGMA foreign_keys = OFF")

    connection.execute("BEGIN IMMEDIATE")
    try:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS schema_migrations (
                version INTEGER PRIMARY KEY,
                applied_at TEXT NOT NULL DEFAULT (datetime('now'))
            )
            """
        )
        applied = {
            row[0] for row in connection.execute("SELECT version FROM schema_migrations")
        }
        for version in sorted(MIGRATIONS):
            if version in applied:
                continue
            MIGRATIONS[version](connection)
            connection.execute(
                "INSERT INTO schema_migrations(version) VALUES (?)", (version,)
            )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        if rebuild_review_issues and foreign_keys_enabled:
            connection.execute("PRAGMA foreign_keys = ON")

    row = connection.execute("SELECT MAX(version) FROM schema_migrations").fetchone()
    return int(row[0] or 0)
