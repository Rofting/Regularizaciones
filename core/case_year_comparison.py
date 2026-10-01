"""Comparación conservadora con el ejercicio previo, sin modificar el cálculo."""
from __future__ import annotations

from calendar import monthrange
from datetime import date
from decimal import Decimal, InvalidOperation


def _day(value):
    return date.fromisoformat(str(value))


def _next_year(value):
    return value.replace(year=value.year + 1,
                         day=min(value.day, monthrange(value.year + 1, value.month)[1]))


def comparable_interval(current, previous, tolerance):
    try:
        start, end = map(_day, current)
        old_start, old_end = map(_day, previous)
    except ValueError:
        return False
    days, old_days = (end - start).days, (old_end - old_start).days
    return (days > 0 and old_days > 0
            and abs(Decimal(days - old_days)) <= Decimal(old_days) * tolerance / 100
            and abs((start - _next_year(old_start)).days) <= 31
            and abs((end - _next_year(old_end)).days) <= 31)


def reference_snapshot(con, case):
    """Incluye las fuentes históricas en las huellas de revisión y de salidas."""
    previous = con.execute('''SELECT * FROM periodos WHERE id_comunidad=?
        AND id_periodo<>? AND fecha_fin IS NOT NULL AND fecha_fin<=?
        ORDER BY fecha_fin DESC,id_periodo DESC LIMIT 1''',
        (case['id_comunidad'], case['id_periodo'], case['fecha_inicio'])).fetchone()
    if previous is None:
        return {}
    period = dict(previous)
    alternatives = [dict(r) for r in con.execute('''SELECT * FROM periodos WHERE id_comunidad=?
        AND id_periodo<>? AND fecha_fin=? ORDER BY id_periodo''',
        (case['id_comunidad'], case['id_periodo'], period['fecha_fin']))]
    run = con.execute('''SELECT * FROM distribution_runs WHERE id_periodo=?
        AND status='completed' ORDER BY id_distribution_run DESC LIMIT 1''',
        (period['id_periodo'],)).fetchone()
    return dict(period=period, alternatives=alternatives, invoices=[dict(r) for r in con.execute(
        'SELECT * FROM facturas WHERE id_comunidad=? AND id_periodo=? ORDER BY id_factura',
        (case['id_comunidad'], period['id_periodo']))],
        readings=[dict(r) for r in con.execute('''SELECT r.* FROM period_readings r
            JOIN propietarios p USING(id_propietario) WHERE p.id_comunidad=?
            AND r.fecha_lectura<=? ORDER BY r.id_propietario,r.tipo,r.fecha_lectura,r.id_periodo''',
            (case['id_comunidad'], period['fecha_fin']))],
        run=dict(run) if run else None,
        owners=[dict(r) for r in con.execute('''SELECT * FROM owner_distribution_snapshots
            WHERE id_distribution_run=? ORDER BY id_snapshot''',
            (run['id_distribution_run'],))] if run else [])


def _amount(value):
    try:
        result = Decimal(str(value))
    except InvalidOperation:
        return None
    return result if result.is_finite() and result >= 0 else None


def _change(current, previous, factor):
    if current == previous:
        return None
    if min(current, previous) == 0:
        return 'pasa de cero a consumo positivo' if previous == 0 else 'pasa de consumo positivo a cero'
    ratio = max(current, previous) / min(current, previous)
    if ratio >= factor:
        return f'variación ×{ratio:.2f} ({"sube" if current > previous else "baja"})'
    return None


def evaluate(con, case, profile, settings, invoices, owners, reference):
    """Devuelve avisos (clave, mensaje) y notas para el panel de coherencia."""
    from case_distribution import _boundary_reading_dates, _consumption_weights, DistributionBlockedError
    from case_coherence import normalise_unit, normalise_cups
    findings, notes = [], []
    if not reference:
        return findings, ['Ejercicio anterior: no hay un período previo de esta comunidad; no se comparan valores históricos.']
    previous = reference['period']
    if len(reference['alternatives']) != 1:
        return findings, ['Ejercicio anterior: varios períodos cierran en la misma fecha; referencia ambigua, no se comparan valores históricos.']
    tolerance = Decimal(settings['historical_duration_tolerance_percent'])
    factor = Decimal(settings['historical_change_factor'])
    if not comparable_interval((case['fecha_inicio'], case['fecha_fin']),
                               (previous['fecha_inicio'], previous['fecha_fin']), tolerance):
        return findings, [f'Ejercicio anterior {previous["nombre"]}: no comparable por duración o fechas; no se prorratean valores.']
    notes.append(f'Ejercicio anterior: {previous["nombre"]} ({previous["fecha_inicio"]} a {previous["fecha_fin"]}); '
                 f'actual {case["fecha_inicio"]} a {case["fecha_fin"]}. Avisos desde ×{factor}, tolerancia de duración {tolerance} %.')
    for invoice in invoices:
        label = f'Factura {invoice["num_factura"] or invoice["id_factura"]} · {invoice["tipo_suministro"]}'
        amount = _amount(invoice['importe_total'])
        if amount is None or amount == 0:
            notes.append(f'{label}: abono, rectificación o importe no positivo; se excluye de la comparación ordinaria de importes.')
            continue
        point, unit = normalise_cups(invoice['cups_o_referencia']), normalise_unit(invoice['unidad_consumo'])
        if not point or not unit:
            notes.append(f'{label}: sin CUPS/referencia o unidad; no hay comparación histórica fiable.')
            continue
        candidates = [r for r in reference['invoices']
            if str(r['tipo_suministro']).upper() == str(invoice['tipo_suministro']).upper()
            and normalise_cups(r['cups_o_referencia']) == point
            and normalise_unit(r['unidad_consumo']) == unit
            and comparable_interval((invoice['fecha_inicio'], invoice['fecha_fin']),
                                    (r['fecha_inicio'], r['fecha_fin']), tolerance)]
        if candidates:
            def distance(document):
                return sum(abs((_day(invoice[key]) - _next_year(_day(document[key]))).days)
                           for key in ('fecha_inicio', 'fecha_fin'))
            closest = min(map(distance, candidates))
            candidates = [r for r in candidates if distance(r) == closest]
        # No emparejar arbitrariamente dos facturas cuando una rectifica a otra.
        if len(candidates) != 1:
            notes.append(f'{label}: {len(candidates)} facturas comparables en {previous["nombre"]}; '
                         'no se elige una referencia ambigua ni se inventa una ausente. Revisa posibles rectificaciones.')
            continue
        old = candidates[0]
        for document in (invoice, old):
            if 'rectific' in str(document['notas'] or '').lower():
                notes.append(f'Factura {document["num_factura"] or document["id_factura"]}: anotada como rectificación; revisa su relación con la original.')
        old_amount = _amount(old['importe_total'])
        if old_amount is None or old_amount == 0:
            notes.append(f'{label}: la referencia es un abono, rectificación o importe no positivo; no se compara.')
            continue
        detail = (f'{label}, CUPS/referencia {point}: actual {amount} € '
                  f'({invoice["fecha_inicio"]} a {invoice["fecha_fin"]}); anterior '
                  f'{old["num_factura"] or old["id_factura"]}: {old_amount} € '
                  f'({old["fecha_inicio"]} a {old["fecha_fin"]}).')
        change = _change(amount, old_amount, factor)
        if change:
            findings.append((f'historical_invoice:{invoice["id_factura"]}:{old["id_factura"]}', detail + ' ' + change + '. Revisa tarifas, consumos y rectificaciones.'))
        else:
            notes.append(detail + ' Sin variación que alcance el umbral.')

    old_case = dict(previous)
    for kind, prefix in (('ACS', 'acs_'), ('CALEFACCION', 'heating_')):
        concept = next((c for c in profile.concepts if c.allocation_method == 'consumption' and c.key.startswith(prefix)), None)
        if concept is None:
            continue
        unit = settings['historical_reading_units'][kind]
        if not unit:
            notes.append(f'{kind}: confirma la misma unidad de contador en ambos ejercicios para comparar consumos históricos.')
            continue
        try:
            current_dates = _boundary_reading_dates(con, case, kind)
            previous_dates = _boundary_reading_dates(con, old_case, kind)
        except DistributionBlockedError as error:
            notes.append(f'{kind}: sin lecturas históricas comparables: {error}.')
            continue
        if not comparable_interval(current_dates, previous_dates, tolerance):
            notes.append(f'{kind}: los intervalos reales de lecturas no son comparables; no se prorratea.')
            continue
        for owner in owners:
            code = owner['codigo_vivienda']
            old_snapshots = [r for r in reference['owners'] if r['id_propietario'] == owner['id_propietario']]
            old_codes = {r['dwelling_code'] for r in old_snapshots if r['dwelling_code']}
            if old_codes and old_codes != {code}:
                notes.append(f'Vivienda {code} · {kind}: el código histórico era {", ".join(sorted(old_codes))}; identidad de vivienda pendiente de revisar, no se compara.')
                continue
            recorded_names = {r['owner_name'] for r in old_snapshots if r['owner_name']}
            if recorded_names and owner['nombre_propietario'] not in recorded_names:
                notes.append(f'Vivienda {code}: cambio de titular registrado; anterior {", ".join(sorted(recorded_names))}; actual {owner["nombre_propietario"]}. Se compara la vivienda, no la persona.')
            elif not recorded_names:
                notes.append(f'Vivienda {code}: histórico sin titular registrado; no se puede comprobar un cambio de titular.')
            old_units = {normalise_unit(r['consumption_unit']) for r in old_snapshots
                         if r['concept_key'].startswith(prefix) and r['consumption_unit']}
            if old_units and old_units != {unit}:
                notes.append(f'Vivienda {code} · {kind}: la unidad histórica difiere de la declarada; no se compara.')
                continue
            try:
                current = _consumption_weights(con, case, concept, [owner['id_propietario']])[owner['id_propietario']]
                old = _consumption_weights(con, old_case, concept, [owner['id_propietario']])[owner['id_propietario']]
            except (DistributionBlockedError, InvalidOperation, ValueError) as error:
                notes.append(f'Vivienda {code} · {kind}: no comparable: {error}.')
                continue
            if not current.is_finite() or not old.is_finite():
                notes.append(f'Vivienda {code} · {kind}: lectura inválida; no se compara.')
                continue
            detail = (f'Vivienda {code} · {kind}: actual {current} {unit} ({current_dates[0]} a {current_dates[1]}); '
                      f'anterior {old} {unit} ({previous_dates[0]} a {previous_dates[1]}).')
            change = _change(current, old, factor)
            if change:
                findings.append((f'historical_consumption:{owner["id_propietario"]}:{kind}:{previous["id_periodo"]}',
                                 detail + ' ' + change + '. Revisa lecturas, ocupación y cambios de contador.'))
            else:
                notes.append(detail + ' Sin variación que alcance el umbral.')
    return findings, notes
