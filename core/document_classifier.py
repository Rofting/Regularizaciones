"""Deterministic document classification before provider detection.

The classifier deliberately favours explicit exclusion types over generic invoice
signals such as dates and totals.  This prevents quotes, delivery notes and bank
receipts from entering the accounting flow merely because they contain an amount.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass


@dataclass(frozen=True)
class DocumentClassification:
    kind: str
    confidence: str
    rule_id: str
    evidence: tuple[str, ...]


def normalize_detection_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    ascii_text = "".join(char for char in normalized if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", ascii_text.upper()).strip()


def _contains(text: str, *patterns: str) -> tuple[str, ...]:
    return tuple(pattern for pattern in patterns if re.search(pattern, text, re.IGNORECASE))


def _classification(
    kind: str,
    confidence: str,
    rule_suffix: str,
    evidence: tuple[str, ...],
) -> DocumentClassification:
    return DocumentClassification(kind, confidence, f"document:{rule_suffix}:v1", evidence)


def _looks_like_reading(text: str) -> tuple[bool, tuple[str, ...]]:
    table_markers = _contains(
        text,
        r"\bLECTURA(?:S)?\b",
        r"\bLECTURA ANTERIOR\b",
        r"\bLECTURA ACTUAL\b",
        r"\bCONTADOR(?:ES)?\b",
        r"\bCONSUMO\b",
    )
    context_markers = _contains(
        text,
        r"\bPROPIEDAD\b",
        r"\bVIVIENDA\b",
        r"\bACS\b",
        r"\bCALEFACCION\b",
        r"\bM3\b",
        r"\bKWH\b",
        r"\bPA\s*\d+[- ]",
    )
    strong_pair = bool(
        re.search(r"\bLECTURA ANTERIOR\b", text)
        and re.search(r"\bLECTURA ACTUAL\b", text)
    )
    return (strong_pair or (len(table_markers) >= 2 and bool(context_markers))), (
        table_markers + context_markers
    )


def _looks_like_invoice(text: str) -> tuple[bool, tuple[str, ...]]:
    invoice_markers = _contains(
        text,
        r"\bFACTURA\b",
        r"\bFACTURA\s*(?:N[OU]|NUMERO|N[º°])\b",
        r"\bN[º°]\s*FACTURA\b",
    )
    structure_markers = _contains(
        text,
        r"\bBASE IMPONIBLE\b",
        r"\bIVA\b",
        r"\bTOTAL(?:\s+A\s+PAGAR)?\b",
        r"\bFECHA(?:\s+DE)?\s+FACTURA\b",
        r"\bPERIODO\s+(?:DE\s+)?FACTURACION\b",
        r"\bVENCIMIENTO\b",
    )
    return bool(invoice_markers and len(structure_markers) >= 2), (
        invoice_markers + structure_markers
    )


def classify_document(text: str, filename: str) -> DocumentClassification:
    normalized = normalize_detection_text(f"{filename or ''}\n{text or ''}")
    if not normalize_detection_text(text):
        return _classification("unknown", "low", "unknown", ())

    explicit_rules = (
        (
            "credit_note",
            "credit-note",
            (r"\bFACTURA RECTIFICATIVA\b", r"\bNOTA DE CREDITO\b", r"\bABONO\b"),
        ),
        (
            "owners",
            "owners",
            (r"\bLISTADO DE PROPIETARIOS\b", r"\bRELACION DE PROPIETARIOS\b"),
        ),
        (
            "quote",
            "quote",
            (r"\bPRESUPUESTO\b", r"\bOFERTA (?:COMERCIAL|ECONOMICA)\b"),
        ),
        (
            "delivery_note",
            "delivery-note",
            (r"\bALBARAN\b", r"\bNOTA DE ENTREGA\b"),
        ),
        (
            "bank_receipt",
            "bank-receipt",
            (
                r"\bJUSTIFICANTE DE (?:TRANSFERENCIA|PAGO)\b",
                r"\bORDEN DE TRANSFERENCIA\b",
                r"\bADEUDO SEPA\b",
            ),
        ),
        (
            "report",
            "report",
            (r"\bINFORME TECNICO\b", r"\bACTA DE (?:INSPECCION|REVISION)\b"),
        ),
    )
    for kind, rule_suffix, patterns in explicit_rules:
        evidence = _contains(normalized, *patterns)
        if evidence:
            return _classification(kind, "high", rule_suffix, evidence)

    is_reading, reading_evidence = _looks_like_reading(normalized)
    if is_reading:
        return _classification("reading", "high", "reading", reading_evidence)

    is_invoice, invoice_evidence = _looks_like_invoice(normalized)
    if is_invoice:
        return _classification("invoice", "high", "invoice", invoice_evidence)

    probable_invoice = _contains(normalized, r"\bFACTURA\b", r"\bTOTAL\b")
    if len(probable_invoice) == 2:
        return _classification(
            "invoice", "medium", "invoice-probable", probable_invoice
        )

    other_evidence = _contains(
        normalized,
        r"\bPARTE DE TRABAJO\b",
        r"\bCERTIFICADO\b",
        r"\bCONTRATO\b",
        r"\bCOMUNICADO\b",
    )
    if other_evidence:
        return _classification("other", "high", "other", other_evidence)

    return _classification("unknown", "low", "unknown", ())


__all__ = [
    "DocumentClassification",
    "classify_document",
    "normalize_detection_text",
]
