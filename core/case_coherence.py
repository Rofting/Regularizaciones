"""Coherencia del expediente y decisiones ligadas a los datos revisados."""
from __future__ import annotations

import hashlib
import json
import sqlite3
from dataclasses import dataclass
from datetime import UTC, date, datetime
from decimal import Decimal, InvalidOperation


@dataclass(frozen=True)
class Finding:
    key: str
    message: str
    can_accept: bool = True
    accepted: bool = False


@dataclass(frozen=True)
class CoherenceReport:
    signature: str
    findings: tuple[Finding, ...]
    notes: tuple[str, ...]

    @property
    def pending(self):
        return tuple(item for item in self.findings if not item.accepted)


def _case(con, case_id):
    row = con.execute("SELECT * FROM regularization_cases WHERE id_case=?", (case_id,)).fetchone()
    if row is None or row['id_periodo'] is None:
        raise ValueError('El expediente debe tener un período enlazado')
    return row


def _read(con, case, key, default):
    row = con.execute('SELECT text_value FROM period_parameters WHERE id_comunidad=? AND id_periodo=? AND parameter_key=?',
                      (case['id_comunidad'], case['id_periodo'], key)).fetchone()
    if row is None:
        return default
    try:
        return json.loads(row[0])
    except (TypeError, ValueError) as error:
        raise ValueError('No se puede leer la configuración o las decisiones de coherencia') from error


def _write(con, case, key, value):
    con.execute('''INSERT INTO period_parameters(id_comunidad,id_periodo,parameter_key,text_value)
        VALUES (?,?,?,?) ON CONFLICT(id_comunidad,id_periodo,parameter_key)
        DO UPDATE SET text_value=excluded.text_value''',
                (case['id_comunidad'], case['id_periodo'], key, json.dumps(value, ensure_ascii=False)))


def normalise_unit(value):
    return str(value or '').strip().lower().replace('³', '3').replace(' ', '')


def normalise_cups(value):
    return ''.join(str(value or '').upper().split())


def _number(value):
    try:
        return Decimal(str(value).strip().replace('%', '').replace(',', '.'))
    except InvalidOperation as error:
        raise ValueError('Introduce un número válido') from error


def validate_settings(settings):
    if not isinstance(settings, dict):
        raise ValueError('La configuración de coherencia debe ser un objeto')
    mode = settings.get('coefficient_mode', 'weights')
    if mode not in {'weights', 'percent'}:
        raise ValueError('Selecciona pesos relativos o porcentajes')
    tolerance = _number(settings.get('consumption_tolerance_percent', 10))
    if not tolerance.is_finite() or not 0 <= tolerance <= 100:
        raise ValueError('La tolerancia de consumo debe estar entre 0 y 100 %')
    comparisons = settings.get('comparisons', [])
    if not isinstance(comparisons, list):
        raise ValueError('Las comparaciones deben ser una lista')
    cleaned = []
    for rule in comparisons:
        if not isinstance(rule, dict) or rule.get('reading_type') not in {'ACS', 'CALEFACCION'}:
            raise ValueError('La comparación debe indicar ACS o calefacción')
        supply = str(rule.get('invoice_type', '')).strip().upper()
        units = [normalise_unit(rule.get(key)) for key in ('invoice_unit', 'reading_unit')]
        if not supply or not all(units):
            raise ValueError('Indica el suministro y las unidades de facturas y lecturas')
        cleaned.append(dict(invoice_type=supply, reading_type=rule['reading_type'],
                            invoice_unit=units[0], reading_unit=units[1], cups=normalise_cups(rule.get('cups'))))
    return dict(coefficient_mode=mode, consumption_tolerance_percent=str(tolerance), comparisons=cleaned)


def load_settings(con, case_id):
    return validate_settings(_read(con, _case(con, case_id), f'coherence_settings:{case_id}', {}))


def evaluate(con: sqlite3.Connection, case_id: int, profile) -> CoherenceReport:
    case = _case(con, case_id)
    settings = load_settings(con, case_id)
    invoices = con.execute('''SELECT * FROM facturas WHERE id_comunidad=? AND id_periodo=? ORDER BY id_factura''',
                           (case['id_comunidad'], case['id_periodo'])).fetchall()
    owners = con.execute('''SELECT id_propietario,codigo_vivienda,coeficiente FROM propietarios
        WHERE id_comunidad=? AND activo=1 AND tipo_unidad='vivienda' ORDER BY id_propietario''',
                         (case['id_comunidad'],)).fetchall()
    readings = con.execute('''SELECT r.* FROM period_readings r JOIN propietarios p USING(id_propietario)
        WHERE p.id_comunidad=? AND r.fecha_lectura<=? ORDER BY r.id_propietario,r.tipo,r.fecha_lectura,r.id_periodo''',
                           (case['id_comunidad'], case['fecha_fin'])).fetchall()
    snapshot = dict(case=dict(case), invoices=[dict(r) for r in invoices], owners=[dict(r) for r in owners],
                    readings=[dict(r) for r in readings], settings=settings,
                    profile=[profile.key, profile.version, profile.source_sha256,
                             [(c.key, c.allocation_method, c.required) for c in profile.concepts]])
    # El estado cambia al generar salidas; no cambia lo que el usuario revisó.
    snapshot['case'] = {k: case[k] for k in ('id_case', 'id_comunidad', 'id_periodo', 'fecha_inicio', 'fecha_fin')}
    signature = hashlib.sha256(json.dumps(snapshot, sort_keys=True, default=str).encode()).hexdigest()
    reviews = _read(con, case, f'coherence_reviews:{case_id}', [])
    if not isinstance(reviews, list):
        raise ValueError('El registro de decisiones de coherencia no es válido')
    accepted = {r.get('key') for r in reviews if isinstance(r, dict) and r.get('signature') == signature and r.get('reason')}
    findings = []
    notes = []

    def add(key, message, can_accept=True):
        findings.append(Finding(key, message, can_accept, can_accept and key in accepted))

    def uses_coefficients(concept):
        if concept.allocation_method != 'coefficient':
            return False
        if concept.required:
            return True
        for source in (concept.actual_source, concept.billed_source):
            if source and source.startswith('period_parameters.'):
                row = con.execute('SELECT numeric_value FROM period_parameters WHERE id_comunidad=? AND id_periodo=? AND parameter_key=?',
                                  (case['id_comunidad'], case['id_periodo'], source.split('.', 1)[1])).fetchone()
                if row and row[0] is not None and _number(row[0]) != 0:
                    return True
        return False

    if any(uses_coefficients(c) for c in profile.concepts):
        try:
            weights = [_number(r['coeficiente']) for r in owners]
        except ValueError:
            weights = []
        if not weights or any(not v.is_finite() or v <= 0 for v in weights):
            add('coefficients_invalid', 'Hay coeficientes nulos, negativos o inválidos. Corrige el listado de viviendas.', False)
        else:
            total = sum(weights)
            if settings['coefficient_mode'] == 'percent' and abs(total - 100) > Decimal('0.01'):
                add('coefficients_total', f'Los porcentajes suman {total} %, deben sumar 100 % (tolerancia 0,01).', False)
            else:
                notes.append(f'Coeficientes: suma {total}; ' + ('porcentajes comprobados.' if settings['coefficient_mode'] == 'percent'
                             else 'pesos relativos normalizados por su suma para repartir.'))

    dated = []
    for invoice in invoices:
        try:
            start, end = (date.fromisoformat(str(invoice[k])) for k in ('fecha_inicio', 'fecha_fin'))
        except ValueError:
            continue  # El control de campos obligatorios ya señala las fechas ausentes.
        if start >= end:
            add(f'invoice_dates:{invoice["id_factura"]}',
                f'Factura {invoice["num_factura"] or invoice["id_factura"]}: el inicio debe ser anterior al fin.', False)
            continue
        dated.append((invoice, start, end))
    for index, (first, start, end) in enumerate(dated):
        for second, other_start, other_end in dated[index + 1:]:
            if first['tipo_suministro'] != second['tipo_suministro']:
                continue
            a, b = (normalise_cups(r['cups_o_referencia']) for r in (first, second))
            if a and b and a != b:
                continue
            if max(start, other_start) < min(end, other_end):
                certainty = 'Solape' if a and b else 'Posible solape (falta identificar el punto de suministro)'
                add(f'overlap:{first["id_factura"]}:{second["id_factura"]}',
                    f'{certainty}: {first["num_factura"] or first["id_factura"]} y '
                    f'{second["num_factura"] or second["id_factura"]} de {first["tipo_suministro"]}, '
                    f'{max(start, other_start)} a {min(end, other_end)}. Revisa duplicados, abonos o rectificaciones.')

    if not settings['comparisons']:
        notes.append('Consumos no comparados: declara qué suministro y contador son equivalentes y sus unidades. Agua general y ACS pueden medir consumos diferentes.')
    for index, rule in enumerate(settings['comparisons']):
        label = f'{rule["invoice_type"]} → {rule["reading_type"]}'
        selected = [r for r in invoices if str(r['tipo_suministro']).upper() == rule['invoice_type']
                    and (not rule['cups'] or normalise_cups(r['cups_o_referencia']) == rule['cups'])]
        units = {normalise_unit(r['unidad_consumo']) for r in selected}
        points = {normalise_cups(r['cups_o_referencia']) for r in selected}
        if not selected or units != {rule['invoice_unit']} or rule['invoice_unit'] != rule['reading_unit'] or len(points) != 1 or not all(points):
            notes.append(f'{label}: no comparable; faltan facturas, unidades equivalentes o un único punto de suministro identificado.')
            continue
        # Usar las mismas lecturas efectivas y aprobaciones que el reparto.
        from case_distribution import _boundary_reading_dates, _consumption_weights, DistributionBlockedError
        concept = next((c for c in profile.concepts if c.allocation_method == 'consumption'
                        and c.key.startswith('acs_' if rule['reading_type'] == 'ACS' else 'heating_')), None)
        if concept is None:
            notes.append(f'{label}: no hay un concepto de consumo compatible en el perfil.')
            continue
        try:
            initial, final = _boundary_reading_dates(con, case, rule['reading_type'])
            amounts = _consumption_weights(con, case, concept, [r['id_propietario'] for r in owners])
        except DistributionBlockedError as error:
            add(f'consumption_readings:{index}', f'{label}: {error}.', False)
            continue
        intervals = sorted((str(r['fecha_inicio']), str(r['fecha_fin'])) for r in selected)
        if (intervals[0][0], intervals[-1][1]) != (initial, final) or any(a[1] != b[0] for a, b in zip(intervals, intervals[1:])):
            notes.append(f'{label}: no comparable; las facturas no cubren sin huecos ni solapes el intervalo de lecturas {initial} a {final}. No se prorratea el consumo.')
            continue
        try:
            billed_values = [_number(r['consumo_total']) for r in selected if r['consumo_total'] is not None]
        except ValueError:
            billed_values = []
        if len(billed_values) != len(selected) or any(not v.is_finite() or v < 0 for v in billed_values):
            add(f'consumption_invalid:{index}', f'{label}: falta un consumo facturado válido.', False)
            continue
        billed, measured = sum(billed_values), sum(amounts.values())
        difference = abs(billed - measured)
        tolerance = abs(billed) * Decimal(settings['consumption_tolerance_percent']) / 100
        message = f'{label}, {initial} a {final}: facturas {billed} {rule["invoice_unit"]}; viviendas {measured} {rule["reading_unit"]}; diferencia {difference}.'
        if difference > tolerance:
            add(f'consumption_difference:{index}', message + f' Supera la tolerancia del {settings["consumption_tolerance_percent"]} %. Revisa lecturas, usos comunes o pérdidas.')
        else:
            notes.append(message + ' Dentro de la tolerancia.')
    return CoherenceReport(signature, tuple(findings), tuple(notes))


def save_settings(con, case_id, settings, coefficients=None):
    case = _case(con, case_id)
    cleaned = validate_settings(settings)
    owners = {r[0] for r in con.execute("SELECT id_propietario FROM propietarios WHERE id_comunidad=? AND activo=1 AND tipo_unidad='vivienda'", (case['id_comunidad'],))}
    values = {}
    for owner, raw in (coefficients or {}).items():
        amount = _number(raw)
        if owner not in owners or not amount.is_finite() or amount < 0:
            raise ValueError('El coeficiente o la vivienda no son válidos')
        values[owner] = amount
    from database_backup import backup_connection
    backup_connection(con, reason='before_coherence_settings')
    with con:
        for owner, amount in values.items():
            con.execute('UPDATE propietarios SET coeficiente=? WHERE id_propietario=?', (float(amount), owner))
        _write(con, case, f'coherence_settings:{case_id}', cleaned)


def accept_finding(con, case_id, profile, key, reason, actor='usuario_local', *, expected_signature=None):
    reason, actor = str(reason).strip(), str(actor).strip()
    if not reason or not actor:
        raise ValueError('Indica el motivo de la aceptación y quién la realiza')
    report = evaluate(con, case_id, profile)
    if expected_signature is not None and report.signature != expected_signature:
        raise ValueError('Los datos cambiaron desde que abriste la revisión. Actualiza el panel antes de aceptar.')
    finding = next((r for r in report.pending if r.key == key), None)
    if finding is None or not finding.can_accept:
        raise ValueError('Este aviso debe corregirse o ya no corresponde a los datos actuales')
    case = _case(con, case_id)
    history = _read(con, case, f'coherence_reviews:{case_id}', [])
    history.append(dict(key=key, signature=report.signature, reason=reason, actor=actor,
                        reviewed_at=datetime.now(UTC).isoformat()))
    from database_backup import backup_connection
    backup_connection(con, reason='before_coherence_review')
    with con:
        _write(con, case, f'coherence_reviews:{case_id}', history)
