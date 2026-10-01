"""Field-level invoice extraction shared by provider families.

Only labelled values are accepted.  Every value carries its rule, confidence
and local text fragment so low-confidence guesses never silently reach a case.

Amounts are collected as *candidates* and the combination that satisfies
``base + IVA = total`` wins; that arithmetic check is what lets the extractor
skip subtotals, percentages and unrelated figures that share a label.
"""

from __future__ import annotations

import itertools
import re
import unicodedata
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from types import MappingProxyType
from typing import Callable, Iterator, Mapping

from provider_registry import ProviderProfile


EXTRACTOR_VERSION = "invoice-fields-v2"

_MONTHS = {
    "enero": 1, "ene": 1, "febrero": 2, "feb": 2, "marzo": 3, "mar": 3,
    "abril": 4, "abr": 4, "mayo": 5, "may": 5, "junio": 6, "jun": 6,
    "julio": 7, "jul": 7, "agosto": 8, "ago": 8, "septiembre": 9,
    "setiembre": 9, "sept": 9, "sep": 9, "octubre": 10, "oct": 10,
    "noviembre": 11, "nov": 11, "diciembre": 12, "dic": 12,
}
_MONTH_NAMES = "|".join(sorted(_MONTHS, key=len, reverse=True))
_DATE_NUMERIC = r"\d{1,2}\s?[./-]\s?\d{1,2}\s?[./-]\s?(?:\d{4}|\d{2})(?!\d)"
_DATE_ISO = r"\d{4}-\d{2}-\d{2}"
_DATE_TEXT = (
    rf"\d{{1,2}}\s+(?:de\s+)?(?:{_MONTH_NAMES})\.?\s+(?:de\s+)?(?:\d{{4}}|\d{{2}})(?!\d)"
)
_DATE = rf"(?:{_DATE_ISO}|{_DATE_NUMERIC}|{_DATE_TEXT})"
_RANGE_SEPARATOR = r"\s*(?:a|al|hasta(?:\s+el)?|y(?:\s+el)?|-|–|—)\s*"

# 1.234,56 | 1234,56 | 1,234.56 | 1234.56, optionally with a trailing minus.
_MONEY_CORE = (
    r"-?(?:\d{1,3}(?:\.\d{3})+,\d{2}|\d+,\d{2}|\d{1,3}(?:,\d{3})+\.\d{2}|\d+\.\d{2})-?"
)
_MONEY_RE = re.compile(rf"(?<![\d.,])({_MONEY_CORE})(?![\d])(?!\s?%)")
_QUANTITY = r"(?P<value>-?\d{1,3}(?:\.\d{3})+(?:,\d+)?|-?\d+(?:[.,]\d+)?)"
_TOKEN = r"(?=[A-Z0-9/._-]*\d)[A-Z0-9][A-Z0-9/._-]{1,30}"
_CUPS_RE = re.compile(
    r"\bES\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{4}\s?[A-Z]{2}(?:\s?\d\s?[A-Z])?\b",
    re.IGNORECASE,
)
_FLAGS = re.IGNORECASE | re.MULTILINE


@dataclass(frozen=True)
class FieldEvidence:
    value: str
    confidence: str
    source: str
    locator: Mapping[str, object]
    rule_id: str
    extractor_version: str = EXTRACTOR_VERSION

    def __post_init__(self) -> None:
        object.__setattr__(self, "locator", MappingProxyType(dict(self.locator)))


@dataclass(frozen=True)
class ExtractionBundle:
    fields: Mapping[str, FieldEvidence]
    diagnostics: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        object.__setattr__(self, "fields", MappingProxyType(dict(self.fields)))
        object.__setattr__(self, "diagnostics", tuple(self.diagnostics))


@dataclass(frozen=True)
class _Candidate:
    value: Decimal
    fragment: str
    confidence: str = "high"


def _strip_accents(value: str) -> str:
    decomposed = unicodedata.normalize("NFKD", value)
    return "".join(char for char in decomposed if not unicodedata.combining(char))


def _iso_date(value: str) -> str | None:
    """Parse numeric, ISO and "1 de marzo de 2026" dates; reject non-dates."""
    cleaned = _strip_accents(value).strip().lower().rstrip(".")
    parsed: date | None = None
    if re.fullmatch(_DATE_ISO, cleaned):
        try:
            parsed = date.fromisoformat(cleaned)
        except ValueError:
            return None
    else:
        textual = re.fullmatch(
            rf"(\d{{1,2}})\s+(?:de\s+)?({_MONTH_NAMES})\.?\s+(?:de\s+)?(\d{{4}}|\d{{2}})",
            cleaned,
        )
        if textual:
            day, month, year = int(textual[1]), _MONTHS[textual[2]], textual[3]
        else:
            numeric = re.fullmatch(
                r"(\d{1,2})\s?[./-]\s?(\d{1,2})\s?[./-]\s?(\d{4}|\d{2})", cleaned
            )
            if not numeric:
                return None
            day, month, year = int(numeric[1]), int(numeric[2]), numeric[3]
        year_number = int(year) + (2000 if len(year) == 2 else 0)
        try:
            parsed = date(year_number, month, day)
        except ValueError:
            return None
    if not 2000 <= parsed.year <= 2100:
        return None
    return parsed.isoformat()


def _decimal(value: str) -> Decimal | None:
    """Parse Spanish (1.234,56) and English (1,234.56) money, incl. ``12,00-``."""
    cleaned = value.strip().replace(" ", "")
    negative = cleaned.endswith("-") or cleaned.startswith("-")
    cleaned = cleaned.strip("-")
    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        head, _, tail = cleaned.rpartition(",")
        cleaned = f"{head.replace(',', '')}.{tail}" if len(tail) <= 2 else cleaned.replace(",", "")
    try:
        number = Decimal(cleaned)
    except InvalidOperation:
        return None
    return -number if negative else number


def _quantity(value: str) -> str | None:
    cleaned = value.strip().replace(" ", "")
    if "," in cleaned:
        if "." in cleaned:
            cleaned = cleaned.replace(".", "")
        cleaned = cleaned.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(?:\.\d{3})+", cleaned):
        cleaned = cleaned.replace(".", "")
    try:
        number = Decimal(cleaned)
    except InvalidOperation:
        return None
    rendered = format(number, "f")
    return rendered.rstrip("0").rstrip(".") if "." in rendered else rendered


def _evidence(
    field_name: str,
    value: str,
    fragment: re.Match[str] | str | None,
    *,
    confidence: str = "high",
    family: str = "standard_spanish_invoice",
    source: str = "text",
    rule: str = "v1",
) -> FieldEvidence:
    if isinstance(fragment, re.Match):
        fragment = fragment.group(0)
    return FieldEvidence(
        value=value,
        confidence=confidence,
        source=source,
        locator={"fragment": " ".join((fragment or "").split())[:500]},
        rule_id=f"{family}:{field_name}:{rule}",
    )


def _iter_matches(patterns: tuple[str, ...], text: str) -> Iterator[re.Match[str]]:
    for pattern in patterns:
        yield from re.finditer(pattern, text, _FLAGS)


def _first_match(patterns: tuple[str, ...], text: str) -> re.Match[str] | None:
    return next(_iter_matches(patterns, text), None)


# ── Dates ────────────────────────────────────────────────────────────────

_PERIOD_LABEL = (
    r"(?:per[ií]odo(?:\s+de)?(?:\s+(?:facturaci[oó]n|facturado|consumo|lectura|liquidaci[oó]n))?"
    r"|facturaci[oó]n|consumo|servicio|fechas?\s+de\s+(?:consumo|facturaci[oó]n))"
)
_PERIOD_PATTERNS: tuple[tuple[str, str], ...] = (
    (
        rf"\b{_PERIOD_LABEL}\b[^\d\n]{{0,35}}(?P<start>{_DATE}){_RANGE_SEPARATOR}(?P<end>{_DATE})",
        "high",
    ),
    (
        rf"\bdesde(?:\s+el)?\s+(?P<start>{_DATE})\s+(?:hasta|a)(?:\s+el)?\s+(?P<end>{_DATE})",
        "high",
    ),
    (rf"\bentre\s+el\s+(?P<start>{_DATE})\s+y\s+el\s+(?P<end>{_DATE})", "high"),
    (
        rf"\bfecha\s+(?:de\s+)?inicio\b[^\d\n]{{0,15}}(?P<start>{_DATE})"
        rf"[\s\S]{{0,80}}?\bfecha\s+(?:de\s+)?(?:fin|final|t[eé]rmino)\b[^\d\n]{{0,15}}(?P<end>{_DATE})",
        "high",
    ),
    # Unlabelled "del 01/03/2026 al 31/03/2026" is common but could also be a
    # contract term, so it is proposed with lower confidence.
    (rf"\bdel\s+(?P<start>{_DATE})\s+al\s+(?P<end>{_DATE})", "medium"),
)

_INVOICE_DATE_PATTERNS: tuple[tuple[str, str], ...] = (
    (
        rf"\b(?:fecha(?:\s+de)?|emitida(?:\s+el)?)\s+(?:emisi[oó]n|expedici[oó]n|factura(?:ci[oó]n)?)"
        rf"(?:\s+(?:de\s+)?(?:la\s+)?factura)?\s*[:\-]?\s*(?P<value>{_DATE})",
        "high",
    ),
    (rf"\bfactura\b[^\n]{{0,60}}?\b(?:de\s+)?fecha\s*[:\-]?\s*(?P<value>{_DATE})", "high"),
    (rf"\bemitida\s+el\s+(?P<value>{_DATE})", "high"),
    (rf"^\s*fecha\s*[:\-]?\s*(?P<value>{_DATE})", "medium"),
)

_NUMBER_PATTERNS = (
    rf"\bn[uú]mero\s+de\s+factura\s*[:\-#]?\s*(?P<value>{_TOKEN})",
    rf"\bn[º°o.]{{1,2}}\s*(?:de\s*)?(?:factura|fra\.?)\s*[:\-#]?\s*(?P<value>{_TOKEN})",
    rf"\bfactura\s*(?:n[º°o.]{{1,2}}|n[uú]m(?:ero)?\.?)\s*[:\-#]?\s*(?P<value>{_TOKEN})",
    rf"\bfra\.?\s*(?:n[º°o.]{{1,2}})?\s*[:\-#]?\s*(?P<value>{_TOKEN})",
    rf"\bfactura\s*[:\-#]?\s*(?P<value>{_TOKEN})",
)


def _extract_period(text: str, family: str, fields: dict[str, FieldEvidence], diagnostics: list[str]) -> None:
    for pattern, confidence in _PERIOD_PATTERNS:
        for match in re.finditer(pattern, text, _FLAGS):
            start = _iso_date(match.group("start"))
            end = _iso_date(match.group("end"))
            if not start or not end:
                continue
            if end < start:
                diagnostics.append("period_inverted")
                continue
            fields["fecha_inicio"] = _evidence(
                "fecha_inicio", start, match, confidence=confidence, family=family
            )
            fields["fecha_fin"] = _evidence(
                "fecha_fin", end, match, confidence=confidence, family=family
            )
            return


def _extract_invoice_date(text: str, family: str, fields: dict[str, FieldEvidence]) -> None:
    for pattern, confidence in _INVOICE_DATE_PATTERNS:
        for match in re.finditer(pattern, text, _FLAGS):
            value = _iso_date(match.group("value"))
            if value:
                fields["fecha_factura"] = _evidence(
                    "fecha_factura", value, match, confidence=confidence, family=family
                )
                return


def _extract_number(text: str, family: str, fields: dict[str, FieldEvidence]) -> None:
    for match in _iter_matches(_NUMBER_PATTERNS, text):
        value = match.group("value").strip(".,-/_")
        if len(value) < 2 or re.fullmatch(_DATE, value, re.IGNORECASE):
            continue
        fields["num_factura"] = _evidence("num_factura", value, match, family=family)
        return


# ── Amounts ──────────────────────────────────────────────────────────────

_EXCLUDED_TOTAL = (
    r"(?!\s*(?:iva\b|base\b|impuesto|energ|potencia|consumo|t[eé]rmino|periodo|kwh|"
    r"l[ií]nea|cuota|descuento|dto))"
)
_BASE_LABELS = (
    r"base\s+imponible|base\s+imp\.?|importe\s+(?:sin|antes\s+de)\s+(?:iva|impuestos)|"
    r"total\s+(?:sin|antes\s+de)\s+(?:iva|impuestos)|total\s+base|subtotal"
)
_IVA_LABELS = (
    r"(?:cuota\s+(?:de\s+)?)?(?:iva|i\.v\.a\.?)|impuesto\s+sobre\s+el\s+valor\s+a[ñn]adido|"
    r"total\s+iva"
)
_TOTAL_STRONG = (
    r"\btotal\s+(?:a\s+pagar|a\s+abonar|factura|importe(?:\s+factura)?|general|"
    r"(?:\(?iva\s+incluido\)?))\b",
    r"\bimporte\s+(?:total|a\s+pagar|de\s+la\s+factura|factura)\b",
    r"\btotal\s+(?:eur|euros|€)",
    r"\ba\s+pagar\b",
)
_TOTAL_WEAK = rf"\btotal\b{_EXCLUDED_TOTAL}"


def _money_in(segment: str) -> list[tuple[str, Decimal]]:
    found = []
    for match in _MONEY_RE.finditer(segment):
        value = _decimal(match.group(1))
        if value is not None:
            found.append((match.group(1), value))
    return found


def _candidates_after_label(
    label_pattern: str,
    text: str,
    *,
    prefer: str = "first",
    confidence: str = "high",
) -> list[_Candidate]:
    """Money figures following a label on its line or, failing that, the next line."""
    lines = text.splitlines()
    result: list[_Candidate] = []
    for index, line in enumerate(lines):
        for label in re.finditer(label_pattern, line, re.IGNORECASE):
            tail = line[label.end():]
            amounts = _money_in(tail)
            fragment = line[label.start():]
            if not amounts and index + 1 < len(lines):
                nxt = lines[index + 1]
                amounts = _money_in(nxt)
                fragment = f"{line.strip()} {nxt.strip()}"
            ordered = list(reversed(amounts)) if prefer == "last" else amounts
            for _, value in ordered:
                result.append(_Candidate(value, fragment.strip()[:500], confidence))
            break
    return result


def _summary_table_candidates(text: str) -> dict[str, _Candidate]:
    """Read ``Base imponible | Cuota IVA | Total`` header rows with values below."""
    lines = text.splitlines()
    token_re = re.compile(
        r"(?P<base>base\s+imponible|base)|(?P<iva>cuota(?:\s+de)?(?:\s+iva)?|(?<![%\w])iva)|"
        r"(?P<total>total(?:\s+factura)?|importe\s+total)",
        re.IGNORECASE,
    )
    for index, line in enumerate(lines[:-1]):
        if not re.search(r"base", line, re.IGNORECASE) or not re.search(r"\biva\b|cuota", line, re.IGNORECASE):
            continue
        if _money_in(line):
            continue
        labels = []
        for token in token_re.finditer(line):
            name = token.lastgroup
            before = line[max(0, token.start() - 2):token.start()]
            if name == "iva" and "%" in before:
                continue
            if not labels or labels[-1] != name:
                labels.append(name)
        for offset in (1, 2):
            if index + offset >= len(lines):
                break
            values = _money_in(lines[index + offset])
            if len(labels) >= 2 and len(values) == len(labels):
                fragment = f"{line.strip()} {lines[index + offset].strip()}"[:500]
                return {
                    name: _Candidate(values[position][1], fragment, "medium")
                    for position, name in enumerate(labels)
                }
    return {}


def _reconciles(base: Decimal, iva: Decimal, total: Decimal) -> bool:
    return abs(base + iva - total) <= Decimal("0.02")


def _select_amounts(
    base: list[_Candidate], iva: list[_Candidate], total: list[_Candidate]
) -> tuple[dict[str, _Candidate], bool | None]:
    """Choose one candidate per field; ``True`` means base+IVA=total was verified."""
    if base and iva and total:
        for b, i, t in itertools.product(base[:6], iva[:6], total[:6]):
            if _reconciles(b.value, i.value, t.value):
                return {"base_imponible": b, "iva": i, "importe_total": t}, True
        return {
            "base_imponible": base[0],
            "iva": iva[0],
            "importe_total": total[0],
        }, False
    chosen = {}
    for name, items in (("base_imponible", base), ("iva", iva), ("importe_total", total)):
        if items:
            chosen[name] = items[0]
    return chosen, None


def _extract_amounts(
    text: str, family: str, fields: dict[str, FieldEvidence], diagnostics: list[str]
) -> None:
    table = _summary_table_candidates(text)
    base = _candidates_after_label(_BASE_LABELS, text)
    iva = _candidates_after_label(rf"\b(?:{_IVA_LABELS})\b", text, prefer="last")
    total: list[_Candidate] = []
    for pattern in _TOTAL_STRONG:
        total.extend(_candidates_after_label(pattern, text, prefer="last"))
    weak = _candidates_after_label(_TOTAL_WEAK, text, prefer="last", confidence="medium")
    total.extend(item for item in weak if item not in total)
    # Table figures go first: a header row followed by its values is the most
    # positional evidence available.
    if table.get("base"):
        base.insert(0, table["base"])
    if table.get("iva"):
        iva.insert(0, table["iva"])
    if table.get("total"):
        total.insert(0, table["total"])

    chosen, verified = _select_amounts(base, iva, total)
    for name, candidate in chosen.items():
        confidence = candidate.confidence
        rule = "v1"
        if name == "importe_total":
            if verified is True:
                confidence, rule = "high", "reconciled:v1"
            elif verified is False:
                confidence, rule = "low", "not-reconciled:v1"
                diagnostics.append("total_not_reconciled")
            elif candidate.confidence == "high":
                confidence = "medium"
        elif verified is True:
            confidence = "high"
        elif verified is False:
            confidence = "medium"
        fields[name] = _evidence(
            name,
            format(candidate.value, ".2f"),
            candidate.fragment,
            confidence=confidence,
            family=family,
            rule=rule,
        )


# ── Service type, CUPS ───────────────────────────────────────────────────

_SERVICE_KEYWORDS: tuple[tuple[str, str], ...] = (
    ("ELECTRICIDAD", r"\b(?:electricidad|energ[ií]a\s+el[eé]ctrica|t[eé]rmino\s+de\s+potencia|kwh\s+el[eé]ctric)"),
    ("GASOLEO", r"\bgas[oó]leo\b"),
    ("GAS", r"\b(?:gas\s+natural|suministro\s+de\s+gas|tur\s+gas|t[eé]rmino\s+fijo\s+gas|hidrocarburos)\b"),
    ("AGUA", r"\b(?:suministro\s+de\s+agua|abastecimiento|alcantarillado|saneamiento|canon\s+del?\s+agua|consumo\s+de\s+agua)\b"),
    ("ASCENSORES", r"\b(?:ascensor(?:es)?|elevador(?:es)?)\b"),
    ("LIMPIEZA", r"\blimpieza\b"),
    ("MANTENIMIENTO", r"\bmantenimiento\b"),
)


def infer_service_family(text: str) -> tuple[str, str] | None:
    """Guess the supply type from keywords when no provider profile applies."""
    folded = _strip_accents(text or "")
    found: dict[str, list[str]] = {}
    for family, pattern in _SERVICE_KEYWORDS:
        hits = re.findall(_strip_accents(pattern), folded, re.IGNORECASE)
        if hits:
            found.setdefault(family, []).extend(hits)
    # Palabras clave configurables (producto + despacho) en config/palabras_clave.json.
    import keywords
    padded = f" {keywords.fold(text)} "
    for family, words in keywords.section("servicios").items():
        for word in words:
            occurrences = padded.count(f" {word} ")
            if occurrences:
                found.setdefault(family, []).extend([word] * occurrences)
    counts = [(len(hits), family, hits[0]) for family, hits in found.items()]
    if not counts:
        return None
    counts.sort(key=lambda item: -item[0])
    # Two services with the same weight are ambiguous: better to ask.
    if len(counts) > 1 and counts[0][0] == counts[1][0]:
        return None
    return counts[0][1], counts[0][2]


def _extract_cups(text: str, family: str, fields: dict[str, FieldEvidence]) -> None:
    match = _CUPS_RE.search(text)
    if match:
        value = re.sub(r"\s", "", match.group(0)).upper()
        fields["cups"] = _evidence("cups", value, match, family=family)


def _extract_common(profile: ProviderProfile, text: str) -> ExtractionBundle:
    family = profile.extractor_family
    fields: dict[str, FieldEvidence] = {}
    diagnostics: list[str] = []

    _extract_period(text, family, fields, diagnostics)
    _extract_invoice_date(text, family, fields)
    _extract_number(text, family, fields)
    _extract_amounts(text, family, fields, diagnostics)

    if profile.service_family:
        fields["tipo_suministro"] = _evidence(
            "tipo_suministro",
            profile.service_family,
            None,
            family=family,
            source="provider_profile",
        )
    else:
        inferred = infer_service_family(text)
        if inferred:
            fields["tipo_suministro"] = _evidence(
                "tipo_suministro", inferred[0], inferred[1],
                confidence="medium", family=family, source="text_keywords",
            )
    return ExtractionBundle(fields, tuple(diagnostics))


def _add_quantity(
    bundle: ExtractionBundle,
    text: str,
    profile: ProviderProfile,
    field_name: str,
    unit_pattern: str,
) -> ExtractionBundle:
    match = _first_match(
        (
            rf"\b(?:consumo(?:\s+(?:total|facturado|real|medido|registrado))?|"
            rf"energ[ií]a\s+consumida|volumen(?:\s+facturado)?)\b[^\d\n-]{{0,35}}"
            rf"{_QUANTITY}\s*{unit_pattern}(?![\w³])",
        ),
        text,
    )
    if not match:
        return bundle
    value = _quantity(match.group("value"))
    if value is None:
        return bundle
    fields = dict(bundle.fields)
    fields[field_name] = _evidence(
        field_name, value, match, family=profile.extractor_family
    )
    return ExtractionBundle(fields, bundle.diagnostics)


def _with_cups(bundle: ExtractionBundle, text: str, profile: ProviderProfile) -> ExtractionBundle:
    fields = dict(bundle.fields)
    _extract_cups(text, profile.extractor_family, fields)
    return ExtractionBundle(fields, bundle.diagnostics)


def _extract_standard(profile: ProviderProfile, text: str) -> ExtractionBundle:
    return _extract_common(profile, text)


def _extract_electricity(profile: ProviderProfile, text: str) -> ExtractionBundle:
    bundle = _add_quantity(_extract_common(profile, text), text, profile, "consumo_kwh", r"kwh")
    return _with_cups(bundle, text, profile)


def _extract_gas_fuel(profile: ProviderProfile, text: str) -> ExtractionBundle:
    result = _add_quantity(_extract_common(profile, text), text, profile, "consumo_kwh", r"kwh")
    result = _add_quantity(result, text, profile, "consumo_m3", r"m(?:3|³)")
    return _with_cups(result, text, profile)


def _extract_water(profile: ProviderProfile, text: str) -> ExtractionBundle:
    return _add_quantity(_extract_common(profile, text), text, profile, "consumo_m3", r"m(?:3|³)")


_FAMILY_EXTRACTORS: Mapping[str, Callable[[ProviderProfile, str], ExtractionBundle]] = {
    "standard_spanish_invoice": _extract_standard,
    "electricity": _extract_electricity,
    "gas_fuel": _extract_gas_fuel,
    "periodic_maintenance": _extract_standard,
    "elevators": _extract_standard,
    "boilers_hvac": _extract_standard,
    "metering_management": _extract_standard,
    "water_public_fees": _extract_water,
}


def extract_invoice_fields(profile: ProviderProfile, text: str) -> ExtractionBundle:
    extractor = _FAMILY_EXTRACTORS.get(profile.extractor_family, _extract_standard)
    return extractor(profile, text or "")


__all__ = [
    "EXTRACTOR_VERSION",
    "ExtractionBundle",
    "FieldEvidence",
    "extract_invoice_fields",
    "infer_service_family",
]
