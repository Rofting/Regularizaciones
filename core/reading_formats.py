"""Mapeos de cabeceras aprendidos de informes confirmados."""
import hashlib
import json
import re
import sqlite3
from datetime import date, datetime

import keywords

EXTRACTOR_VERSION = "reading-tables-v2"


def header_signature(cells):
    headers = []
    for cell in cells:
        if isinstance(cell, (date, datetime)):
            headers.append("fecha")
            continue
        text = keywords.fold(cell)
        text = re.sub(r"\b\d{1,2} \d{1,2} \d{2,4}\b", "fecha", text)
        text = re.sub(r"\b(?:ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[a-z]* \d{2,4}\b", "fecha", text)
        text = re.sub(r"^(lectura |lect )?\d{1,2} \d{4}$", r"\1fecha", text)
        headers.append(text)
    return hashlib.sha256(json.dumps(headers, ensure_ascii=False).encode()).hexdigest()


def formats_for(connection, company):
    if connection is None or not company:
        return {}
    try:
        rows = connection.execute(
            "SELECT header_signature,mapping_json FROM learned_reading_formats "
            "WHERE company=? AND extractor_version=?", (company, EXTRACTOR_VERSION),
        ).fetchall()
    except sqlite3.OperationalError:
        return {}
    result = {}
    for signature, raw in rows:
        try:
            mapping = json.loads(raw)
        except (ValueError, TypeError):
            continue
        if _valid_mapping(mapping):
            result[signature] = mapping
    return result


def _valid_mapping(mapping):
    if not isinstance(mapping, list) or not mapping:
        return False
    roles = set(keywords.section("cabeceras_lecturas")) | {"lectura_fecha", "lectura_servicio"}
    valid = all(
        isinstance(item, dict) and type(item.get("index")) is int and item["index"] >= 0
        and isinstance(item.get("role"), str) and item["role"] in roles
        and item.get("service") in (None, "ACS", "CALEFACCION")
        for item in mapping
    )
    return valid and len({item["index"] for item in mapping}) == len(mapping) and any(
        item["role"] == "vivienda" for item in mapping
    ) and sum(item["role"] in {"val_ant", "val_act", "lectura_fecha", "lectura_servicio"}
              for item in mapping) >= 2


def learn_document_format(connection, document_id, confirmed_by):
    """Sólo aprende después de una confirmación humana y aplicación correcta."""
    row = connection.execute(
        "SELECT value FROM extraction_candidates WHERE id_document=? "
        "AND field_name='reading.format'", (document_id,),
    ).fetchone()
    if row is None or not row[0]:
        return
    try:
        metadata = json.loads(row[0])
    except (ValueError, TypeError):
        return
    if (not isinstance(metadata, dict) or not isinstance(metadata.get("company"), str)
            or metadata["company"] not in keywords.section("empresas_lecturas")
            or metadata.get("extractor_version") != EXTRACTOR_VERSION
            or not isinstance(metadata.get("header_signature"), str)
            or not re.fullmatch(r"[a-f0-9]{64}", metadata["header_signature"])
            or not _valid_mapping(metadata.get("mapping"))):
        return
    connection.execute(
        """INSERT INTO learned_reading_formats
           (company,header_signature,extractor_version,mapping_json,id_document,confirmed_by)
           VALUES (?,?,?,?,?,?)
           ON CONFLICT(company,header_signature,extractor_version) DO UPDATE SET
           mapping_json=excluded.mapping_json,id_document=excluded.id_document,
           confirmed_by=excluded.confirmed_by,learned_at=datetime('now')""",
        (metadata["company"], metadata["header_signature"], EXTRACTOR_VERSION,
         json.dumps(metadata["mapping"], ensure_ascii=False), document_id, confirmed_by),
    )
