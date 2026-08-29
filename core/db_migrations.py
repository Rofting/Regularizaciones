"""Migraciones incrementales y transaccionales de la base de datos."""

from __future__ import annotations

import sqlite3
from collections.abc import Callable


CURRENT_SCHEMA_VERSION = 1


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


MIGRATIONS: dict[int, Callable[[sqlite3.Connection], None]] = {1: _migration_1}


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
