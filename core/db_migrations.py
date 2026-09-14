"""Migraciones incrementales y transaccionales de la base de datos."""

from __future__ import annotations

import sqlite3
from collections.abc import Callable


CURRENT_SCHEMA_VERSION = 8


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


MIGRATIONS: dict[int, Callable[[sqlite3.Connection], None]] = {
    1: _migration_1,
    2: _migration_2,
    3: _migration_3,
    4: _migration_4,
    5: _migration_5,
    6: _migration_6,
    7: _migration_7,
    8: _migration_8,
}


def migrate(connection: sqlite3.Connection) -> int:
    """Aplica en orden las migraciones pendientes o revierte todo el lote."""
    if connection.in_transaction:
        connection.commit()

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

    row = connection.execute("SELECT MAX(version) FROM schema_migrations").fetchone()
    return int(row[0] or 0)
