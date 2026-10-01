import sys
import unittest
from pathlib import Path
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'core'))
import case_coherence
import excel_export_service
from excel_profiles import load_profile
import tests.test_excel_export_service as fixtures


class PreviousYearComparisonTest(unittest.TestCase):
    _insert_normalized_inputs = fixtures.ExcelExportServiceTest._insert_normalized_inputs
    tearDown = fixtures.ExcelExportServiceTest.tearDown

    def setUp(self):
        fixtures.ExcelExportServiceTest.setUp(self)
        self.profile = load_profile('658_acs_v1', self.project_root)
        self.owners = self.connection.execute('SELECT * FROM propietarios ORDER BY id_propietario').fetchall()

    def report(self):
        return case_coherence.evaluate(self.connection, self.case_id, self.profile)

    def historical(self):
        return [r for r in self.report().pending if r.key.startswith('historical_')]

    def previous(self, consumptions=(25, 15), gas_amount=100):
        self.previous_id = self.connection.execute('''INSERT INTO periodos
            (id_comunidad,nombre,fecha_inicio,fecha_fin) VALUES (?,'2024-2025','2024-09-01','2025-08-31')''',
            (self.community_id,)).lastrowid
        for owner, consumption in zip(self.owners, consumptions):
            first = 100 if owner['codigo_vivienda'] == 'P1-A' else 200
            self.connection.executemany('''INSERT INTO lecturas_vecino
                (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
                VALUES (?,?,'ACS',?,?,'real')''',
                [(owner['id_propietario'], self.previous_id, '2024-09-01', first - consumption),
                 (owner['id_propietario'], self.previous_id, '2025-08-31', first)])
        self.old_invoice = self.connection.execute('''INSERT INTO facturas
            (id_comunidad,id_periodo,tipo_suministro,cups_o_referencia,num_factura,
             fecha_inicio,fecha_fin,importe_total,unidad_consumo)
            VALUES (?,?,'GAS','SINTETICO','G-ANT','2024-09-01','2024-09-30',?,'kWh')''',
            (self.community_id, self.previous_id, gas_amount)).lastrowid
        self.connection.commit()

    def settings(self, **changes):
        settings = case_coherence.load_settings(self.connection, self.case_id)
        settings.update(changes)
        case_coherence.save_settings(self.connection, self.case_id, settings)

    def hash(self):
        return excel_export_service.calculate_case_input_hash(self.connection, id_case=self.case_id, project_root=self.project_root)

    def test_no_history_is_reported_without_inventing_a_baseline(self):
        self.assertEqual([], self.historical())
        self.assertTrue(any('no hay un período previo' in n for n in self.report().notes))

    def test_equal_consumption_and_amounts_do_not_open_warnings(self):
        self.previous()
        self.assertEqual([], self.historical())
        self.assertTrue(any('actual 25.0 m3' in n and 'anterior 25.0 m3' in n for n in self.report().notes))

    def test_invoice_threefold_increase_blocks_direct_excel(self):
        self.previous(gas_amount=25)
        finding = self.historical()[0]
        self.assertTrue(finding.key.startswith('historical_invoice:'))
        self.assertIn('actual 100.0 €', finding.message)
        self.assertIn('25.0 €', finding.message)
        self.assertIn('×4.00', finding.message)
        engine = Mock()
        with self.assertRaisesRegex(excel_export_service.ExportBlockedError, 'G-ANT'):
            excel_export_service.generate_official_excel(self.connection, id_case=self.case_id,
                project_root=self.project_root, output_root=self.output_root, recalculator=engine)
        engine.recalculate.assert_not_called()

    def test_invoice_fall_is_also_detected(self):
        self.previous(gas_amount=300)
        self.assertIn('(baja)', self.historical()[0].message)
        self.assertIn('×3.00', self.historical()[0].message)

    def test_dwelling_consumption_is_compared_not_meter_odometer(self):
        self.previous(consumptions=(5, 15))
        finding = self.historical()[0]
        self.assertIn('Vivienda P1-A', finding.message)
        self.assertIn('actual 25.0 m3', finding.message)
        self.assertIn('anterior 5.0 m3', finding.message)
        self.assertIn('×5.00', finding.message)

    def test_zero_to_positive_and_positive_to_zero_are_detected(self):
        self.previous(consumptions=(0, 15))
        self.assertIn('cero a consumo positivo', self.historical()[0].message)
        self.connection.execute("UPDATE lecturas_vecino SET valor_acumulado=200 WHERE fecha_lectura='2026-08-31' AND id_propietario=?", (self.owners[1]['id_propietario'],))
        self.assertTrue(any('consumo positivo a cero' in f.message for f in self.historical()))

    def test_configurable_factor_suppresses_small_changes_and_validates_input(self):
        self.previous(gas_amount=25)
        self.settings(historical_change_factor='5')
        self.assertEqual([], self.historical())
        for factor in ('1', 'NaN', '-3', 'mal'):
            with self.subTest(factor=factor), self.assertRaises(ValueError):
                self.settings(historical_change_factor=factor)
        self.assertEqual('5', case_coherence.load_settings(self.connection, self.case_id)['historical_change_factor'])

    def test_reference_with_a_shorter_or_shifted_season_is_not_prorated(self):
        self.previous(gas_amount=25, consumptions=(5, 15))
        self.connection.execute("UPDATE periodos SET fecha_inicio='2025-03-01' WHERE id_periodo=?", (self.previous_id,))
        self.assertEqual([], self.historical())
        self.assertTrue(any('no comparable por duración' in n for n in self.report().notes))

    def test_shorter_invoice_interval_is_not_used_even_in_comparable_year(self):
        self.previous(gas_amount=25)
        self.connection.execute("UPDATE facturas SET fecha_inicio='2024-09-20' WHERE id_factura=?", (self.old_invoice,))
        self.assertEqual([], self.historical())

    def test_different_supply_point_or_unit_is_not_used(self):
        self.previous(gas_amount=25)
        for column, value in (('cups_o_referencia', 'OTRO'), ('unidad_consumo', 'm3')):
            with self.subTest(column=column):
                self.connection.execute("UPDATE facturas SET cups_o_referencia='SINTETICO',unidad_consumo='kWh' WHERE id_factura=?", (self.old_invoice,))
                self.connection.execute(f'UPDATE facturas SET {column}=? WHERE id_factura=?', (value, self.old_invoice))
                self.assertEqual([], self.historical())

    def test_ambiguous_previous_invoices_do_not_choose_an_arbitrary_amount(self):
        self.previous(gas_amount=25)
        self.connection.execute('''INSERT INTO facturas(id_comunidad,id_periodo,tipo_suministro,
            cups_o_referencia,num_factura,fecha_inicio,fecha_fin,importe_total,unidad_consumo)
            VALUES (?,?,'GAS','SINTETICO','RECTIFICACION','2024-09-01','2024-09-30',-25,'kWh')''',
            (self.community_id, self.previous_id))
        self.assertEqual([], self.historical())
        self.assertTrue(any('2 facturas comparables' in n for n in self.report().notes))

    def test_adjacent_months_use_the_closest_annual_interval(self):
        self.previous(gas_amount=25)
        self.connection.execute('''INSERT INTO facturas(id_comunidad,id_periodo,tipo_suministro,
            cups_o_referencia,num_factura,fecha_inicio,fecha_fin,importe_total,unidad_consumo)
            VALUES (?,?,'GAS','SINTETICO','OCTUBRE','2024-10-01','2024-10-30',100,'kWh')''',
            (self.community_id, self.previous_id))
        self.assertEqual(1, len(self.historical()))
        self.assertIn('G-ANT', self.historical()[0].message)

    def test_credit_and_annotated_rectification_are_identified(self):
        self.previous(gas_amount=-25)
        self.assertEqual([], self.historical())
        self.assertTrue(any('referencia es un abono' in n for n in self.report().notes))
        self.connection.execute("UPDATE facturas SET importe_total=25,notas='Factura rectificativa' WHERE id_factura=?", (self.old_invoice,))
        self.assertTrue(any('anotada como rectificación' in n for n in self.report().notes))

    def test_approved_reason_is_invalidated_by_a_changed_historical_invoice(self):
        self.previous(gas_amount=25)
        report = self.report()
        old_hash = self.hash()
        case_coherence.accept_finding(self.connection, self.case_id, self.profile, self.historical()[0].key,
                                     'Cambio de tarifa comprobado', expected_signature=report.signature)
        self.assertEqual([], self.historical())
        self.connection.execute('UPDATE facturas SET importe_total=20 WHERE id_factura=?', (self.old_invoice,))
        self.connection.commit()
        self.assertTrue(self.historical())
        self.assertNotEqual(old_hash, self.hash())
        with self.assertRaisesRegex(ValueError, 'datos cambiaron'):
            case_coherence.accept_finding(self.connection, self.case_id, self.profile, self.historical()[0].key,
                'Motivo anterior', expected_signature=report.signature)

    def test_previous_reading_change_also_invalidates_output_identity(self):
        self.previous()
        old_hash = self.hash()
        self.connection.execute("UPDATE lecturas_vecino SET valor_acumulado=98 WHERE fecha_lectura='2024-09-01' AND id_propietario=?", (self.owners[0]['id_propietario'],))
        self.connection.commit()
        self.assertNotEqual(old_hash, self.hash())
        self.assertTrue(self.historical())

    def test_missing_or_unapproved_previous_readings_are_explained(self):
        self.previous()
        self.connection.execute("UPDATE lecturas_vecino SET estado='estimado',metodo_estimacion='manual' WHERE id_periodo=? AND id_propietario=?", (self.previous_id, self.owners[0]['id_propietario']))
        self.assertEqual([], self.historical())
        self.assertTrue(any('P1-A' in n and 'sin aprobación' in n for n in self.report().notes))
        self.connection.execute('DELETE FROM lecturas_vecino WHERE id_periodo=? AND id_propietario=?', (self.previous_id, self.owners[0]['id_propietario']))
        self.assertTrue(any('P1-A' in n and 'Falta una lectura' in n for n in self.report().notes))

    def test_unconfirmed_units_disable_reading_comparison(self):
        self.previous(consumptions=(5, 15))
        self.settings(historical_reading_units={'ACS': '', 'CALEFACCION': ''})
        self.assertEqual([], self.historical())
        self.assertTrue(any('confirma la misma unidad' in n for n in self.report().notes))

    def test_heating_requires_confirmed_units_and_uses_them_in_warnings(self):
        from dataclasses import replace
        from excel_profiles import ConceptRule
        self.previous()
        self.profile = replace(self.profile, concepts=self.profile.concepts + (
            ConceptRule('heating_variable', 'consumption', 'period_parameters.heating_actual', None, True),))
        for owner in self.owners:
            self.connection.executemany('''INSERT INTO lecturas_vecino
                (id_propietario,id_periodo,tipo,fecha_lectura,valor_acumulado,estado)
                VALUES (?,?,'CALEFACCION',?,?,'real')''',
                [(owner['id_propietario'], self.previous_id, '2024-09-01', 0),
                 (owner['id_propietario'], self.previous_id, '2025-08-31', 10),
                 (owner['id_propietario'], self.period_id, '2025-09-01', 10),
                 (owner['id_propietario'], self.period_id, '2026-08-31', 50)])
        self.assertEqual([], self.historical())
        self.assertTrue(any('CALEFACCION: confirma' in n for n in self.report().notes))
        self.settings(historical_reading_units={'ACS': 'm3', 'CALEFACCION': 'unidades'})
        self.assertEqual(2, len(self.historical()))
        self.assertTrue(all('40.0 unidades' in f.message for f in self.historical()))

    def test_carry_forward_does_not_count_a_zero_reading_as_new_consumption(self):
        self.previous(consumptions=(0, 15))
        self.connection.execute("UPDATE lecturas_vecino SET valor_acumulado=0,estado='estimado',metodo_estimacion='carry_forward_zero' WHERE id_periodo=? AND id_propietario=? AND fecha_lectura='2026-08-31'", (self.period_id, self.owners[0]['id_propietario']))
        self.assertEqual([], self.historical())
        self.assertTrue(any('actual 0.0 m3' in n and 'anterior 0.0 m3' in n for n in self.report().notes))

    def test_duration_tolerance_is_configurable_and_rejects_invalid_values(self):
        self.previous(gas_amount=25)
        self.connection.execute("UPDATE periodos SET fecha_inicio='2024-08-31' WHERE id_periodo=?", (self.previous_id,))
        self.settings(historical_duration_tolerance_percent='0')
        self.assertEqual([], self.historical())
        self.settings(historical_duration_tolerance_percent='10')
        self.assertTrue(self.historical())
        for value in ('NaN', '101', '-1'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                self.settings(historical_duration_tolerance_percent=value)

    def test_anniversary_comparison_accepts_a_leap_day_without_changing_season(self):
        from decimal import Decimal
        from case_year_comparison import comparable_interval
        self.assertTrue(comparable_interval(('2025-02-28', '2026-02-28'),
                                            ('2024-02-29', '2025-02-28'), Decimal(10)))
        self.assertFalse(comparable_interval(('2025-07-01', '2026-07-01'),
                                             ('2024-02-29', '2025-02-28'), Decimal(10)))

    def test_multiple_previous_periods_are_reported_as_ambiguous(self):
        self.previous(gas_amount=25)
        self.connection.execute('''INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin)
            VALUES (?,'Alternativo','2024-09-01','2025-08-31')''', (self.community_id,))
        self.assertEqual([], self.historical())
        self.assertTrue(any('referencia ambigua' in n for n in self.report().notes))

    def test_historical_owner_identity_is_not_invented_and_changes_are_shown(self):
        self.previous()
        self.assertTrue(any('histórico sin titular registrado' in n for n in self.report().notes))
        old_case = self.connection.execute('''INSERT INTO regularization_cases
            (id_comunidad,id_periodo,nombre,fecha_inicio,fecha_fin,estado)
            VALUES (?,?,'Anterior','2024-09-01','2025-08-31','calculated')''',
            (self.community_id, self.previous_id)).lastrowid
        run = self.connection.execute('''INSERT INTO distribution_runs
            (id_case,id_periodo,input_sha256,status) VALUES (?,?,'previo','completed')''',
            (old_case, self.previous_id)).lastrowid
        self.connection.execute('''INSERT INTO owner_distribution_snapshots
            (id_distribution_run,id_propietario,concept_key,billed_cents,actual_cents,difference_cents,
             owner_name,dwelling_code,consumption_unit) VALUES (?,?,'acs_variable',0,0,0,'Titular anterior','P1-A','m³')''',
            (run, self.owners[0]['id_propietario']))
        self.assertTrue(any('cambio de titular registrado' in n and 'Titular anterior' in n for n in self.report().notes))
        self.connection.execute("UPDATE owner_distribution_snapshots SET consumption_unit='kWh'")
        self.assertTrue(any('unidad histórica difiere' in n for n in self.report().notes))
        self.connection.execute("UPDATE owner_distribution_snapshots SET dwelling_code='OTRA VIVIENDA'")
        self.assertTrue(any('identidad de vivienda pendiente' in n for n in self.report().notes))

    def test_read_only_report_and_other_community_history_is_ignored(self):
        community = self.connection.execute("INSERT INTO comunidades(codigo,nombre) VALUES ('OTRA','Otra')").lastrowid
        self.connection.execute('''INSERT INTO periodos(id_comunidad,nombre,fecha_inicio,fecha_fin)
            VALUES (?,'2024-2025','2024-09-01','2025-08-31')''', (community,))
        before = self.connection.total_changes
        self.assertTrue(any('no hay un período previo' in n for n in self.report().notes))
        self.assertEqual(before, self.connection.total_changes)


class GroupedNotesTest(unittest.TestCase):
    def test_large_communities_get_one_summary_line(self):
        import case_year_comparison
        many = [f'V{i}' for i in range(30)]
        notes = case_year_comparison._grouped_notes(
            many, [(code, 'ACS', f'{code} sin variación') for code in many])
        self.assertEqual(2, len(notes))
        self.assertIn('30 viviendas', notes[0])
        self.assertIn('ACS: 30 viviendas sin variación', notes[1])
        few = case_year_comparison._grouped_notes(['V1'], [('V1', 'ACS', 'detalle V1')])
        self.assertEqual(['Vivienda V1: histórico sin titular registrado; no se puede comprobar '
                          'un cambio de titular.', 'detalle V1'], few)
