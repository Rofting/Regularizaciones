import json
import sys
import unittest
from pathlib import Path
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'core'))
import case_coherence
import case_readiness
import excel_export_service
from excel_profiles import load_profile
import tests.test_excel_export_service as fixtures


class CaseCoherenceTest(unittest.TestCase):
    _insert_normalized_inputs = fixtures.ExcelExportServiceTest._insert_normalized_inputs
    tearDown = fixtures.ExcelExportServiceTest.tearDown

    def setUp(self):
        fixtures.ExcelExportServiceTest.setUp(self)
        self.profile = load_profile('658_acs_v1', self.project_root)

    def report(self):
        return case_coherence.evaluate(self.connection, self.case_id, self.profile)

    def settings(self, **changes):
        settings = case_coherence.load_settings(self.connection, self.case_id)
        settings.update(changes)
        case_coherence.save_settings(self.connection, self.case_id, settings)

    def comparable(self, consumption=40):
        self.connection.execute("UPDATE facturas SET fecha_inicio='2025-09-01',fecha_fin='2026-08-31',consumo_total=? WHERE tipo_suministro='AGUA'", (consumption,))
        self.connection.commit()
        self.settings(comparisons=[dict(invoice_type='AGUA', reading_type='ACS',
                                       invoice_unit='m³', reading_unit='m3', cups='SINTETICO')])

    def duplicate(self, cups='SINTETICO', amount=100):
        cursor = self.connection.execute('''INSERT INTO facturas(id_comunidad,id_periodo,tipo_suministro,
            cups_o_referencia,num_factura,fecha_inicio,fecha_fin,importe_total)
            VALUES (?,?, 'GAS',?,'G-2','2025-09-15','2025-10-15',?)''',
                                (self.community_id, self.period_id, cups, amount))
        self.connection.execute("UPDATE facturas SET fecha_factura='2025-10-15',dias_facturados=30,consumo_total=50,unidad_consumo='kWh',termino_variable=? WHERE id_factura=?", (amount, cursor.lastrowid))
        self.connection.executemany('INSERT INTO invoice_components(id_factura,component_key,amount,unit) VALUES (?,?,?,?)',
                                   [(cursor.lastrowid, key, value, 'EUR') for key, value in [('fixed', 0), ('variable', amount), ('total', amount)]])
        self.connection.commit()

    def test_overlapping_same_point_invoices_need_review(self):
        self.duplicate()
        self.assertTrue(any(f.key.startswith('overlap:') for f in self.report().pending))

    def test_distinct_supply_points_are_not_duplicates(self):
        self.duplicate(cups='OTRO PUNTO')
        self.assertFalse(any(f.key.startswith('overlap:') for f in self.report().pending))

    def test_unknown_supply_point_marks_possible_overlap(self):
        self.duplicate(cups=None)
        self.assertIn('Posible solape', self.report().pending[0].message)

    def test_shared_boundary_is_not_a_duplicate_day(self):
        self.duplicate()
        self.connection.execute("UPDATE facturas SET fecha_inicio='2025-09-30' WHERE num_factura='G-2'")
        self.assertEqual((), self.report().pending)

    def test_credit_overlap_can_be_accepted_with_an_audited_reason(self):
        self.duplicate(amount=-100)
        finding = self.report().pending[0]
        case_coherence.accept_finding(self.connection, self.case_id, self.profile, finding.key, 'Abono de la factura original')
        self.assertEqual((), self.report().pending)
        row = self.connection.execute("SELECT text_value FROM period_parameters WHERE parameter_key=?",
                                      (f'coherence_reviews:{self.case_id}',)).fetchone()
        record = json.loads(row[0])[0]
        self.assertEqual('Abono de la factura original', record['reason'])
        self.assertTrue(record['reviewed_at'])
        self.assertEqual('usuario_local', record['actor'])

    def test_changed_invoice_invalidates_acceptance(self):
        self.duplicate()
        case_coherence.accept_finding(self.connection, self.case_id, self.profile, self.report().pending[0].key, 'Rectificación verificada')
        self.connection.execute("UPDATE facturas SET importe_total=110 WHERE num_factura='G-2'")
        self.connection.commit()
        self.assertTrue(self.report().pending)

    def test_stale_panel_cannot_accept_new_data_with_an_old_reason(self):
        self.duplicate()
        report = self.report()
        self.connection.execute("UPDATE facturas SET importe_total=110 WHERE num_factura='G-2'")
        self.connection.commit()
        with self.assertRaisesRegex(ValueError, 'datos cambiaron'):
            case_coherence.accept_finding(self.connection, self.case_id, self.profile,
                report.pending[0].key, 'Motivo referido al importe anterior', expected_signature=report.signature)
        self.assertTrue(self.report().pending)

    def test_relative_weights_are_normalised_without_requiring_100(self):
        self.connection.execute('UPDATE propietarios SET coeficiente=3')
        report = self.report()
        self.assertEqual((), report.pending)
        self.assertTrue(any('pesos relativos normalizados' in n for n in report.notes))

    def test_percentages_must_sum_100_and_cannot_be_waived(self):
        self.connection.execute('UPDATE propietarios SET coeficiente=49')
        self.connection.commit()
        self.settings(coefficient_mode='percent')
        finding = self.report().pending[0]
        self.assertEqual('coefficients_total', finding.key)
        with self.assertRaises(ValueError):
            case_coherence.accept_finding(self.connection, self.case_id, self.profile, finding.key, 'Aceptar igualmente')

    def test_correcting_coefficients_clears_percentage_blocker(self):
        owners = [r[0] for r in self.connection.execute('SELECT id_propietario FROM propietarios')]
        case_coherence.save_settings(self.connection, self.case_id, dict(coefficient_mode='percent'),
                                     {owners[0]: '30,0', owners[1]: '70%'})
        self.assertEqual((), self.report().pending)

    def test_same_unit_interval_and_point_are_compared(self):
        self.comparable()
        report = self.report()
        self.assertEqual((), report.pending)
        self.assertTrue(any('Dentro de la tolerancia' in n for n in report.notes))

    def test_large_consumption_difference_needs_review_and_changes_reopen_it(self):
        self.comparable(100)
        finding = self.report().pending[0]
        self.assertIn('viviendas 40', finding.message)
        case_coherence.accept_finding(self.connection, self.case_id, self.profile, finding.key, 'Consumo de zonas comunes revisado')
        self.assertEqual((), self.report().pending)
        self.connection.execute("UPDATE lecturas_vecino SET valor_acumulado=valor_acumulado+1 WHERE fecha_lectura='2026-08-31'")
        self.connection.commit()
        self.assertTrue(self.report().pending)

    def test_different_units_are_not_compared(self):
        self.comparable(100)
        self.settings(comparisons=[dict(invoice_type='AGUA', reading_type='ACS', invoice_unit='m3', reading_unit='kWh')])
        self.assertEqual((), self.report().pending)
        self.assertTrue(any('no comparable' in n for n in self.report().notes))

    def test_invalid_invoice_consumption_requires_correction(self):
        self.comparable()
        self.connection.execute("UPDATE facturas SET consumo_total='ilegible' WHERE tipo_suministro='AGUA'")
        finding = self.report().pending[0]
        self.assertTrue(finding.key.startswith('consumption_invalid:'))
        self.assertFalse(finding.can_accept)

    def test_partial_invoice_period_is_not_prorated(self):
        self.comparable(100)
        self.connection.execute("UPDATE facturas SET fecha_fin='2026-05-31' WHERE tipo_suministro='AGUA'")
        self.assertEqual((), self.report().pending)
        self.assertTrue(any('No se prorratea' in n for n in self.report().notes))

    def test_carry_forward_uses_the_same_consumption_as_distribution(self):
        self.comparable(15)
        self.connection.execute("UPDATE lecturas_vecino SET valor_acumulado=0,estado='estimado',metodo_estimacion='carry_forward_zero' WHERE id_propietario=(SELECT MIN(id_propietario) FROM propietarios) AND fecha_lectura='2026-08-31'")
        report = self.report()
        self.assertEqual((), report.pending)
        self.assertTrue(any('viviendas 15' in n for n in report.notes))

    def test_unapproved_estimation_blocks_comparison(self):
        self.comparable()
        self.connection.execute("UPDATE lecturas_vecino SET estado='estimado',metodo_estimacion='manual' WHERE fecha_lectura='2026-08-31'")
        finding = self.report().pending[0]
        self.assertIn('sin aprobación', finding.message)
        self.assertFalse(finding.can_accept)

    def test_review_does_not_modify_database_and_readiness_routes_to_it(self):
        self.duplicate()
        before = self.connection.total_changes
        self.report()
        report = case_readiness.evaluate_case_readiness(self.connection, self.case_id, self.project_root)
        self.assertTrue(any(b.action == 'review_coherence' for b in report.blockers))
        self.assertEqual(before, self.connection.total_changes)

    def test_direct_export_cannot_bypass_overlap_review(self):
        self.duplicate()
        recalculator = Mock()
        with self.assertRaisesRegex(excel_export_service.ExportBlockedError, 'coherencia'):
            excel_export_service.generate_official_excel(self.connection, id_case=self.case_id,
                project_root=self.project_root, output_root=self.output_root, recalculator=recalculator)
        recalculator.recalculate.assert_not_called()
        self.assertEqual(0, self.connection.execute('SELECT COUNT(*) FROM excel_export_runs').fetchone()[0])

    def test_invalid_configuration_does_not_change_coefficients(self):
        owner = self.connection.execute('SELECT id_propietario FROM propietarios LIMIT 1').fetchone()[0]
        for tolerance in ('NaN', '-1', '101', 'mal'):
            with self.subTest(tolerance=tolerance), self.assertRaises(ValueError):
                case_coherence.save_settings(self.connection, self.case_id,
                    dict(consumption_tolerance_percent=tolerance), {owner: '30'})
        self.assertEqual(50, self.connection.execute('SELECT coeficiente FROM propietarios WHERE id_propietario=?', (owner,)).fetchone()[0])

    def test_empty_reason_cannot_accept_an_overlap(self):
        self.duplicate()
        with self.assertRaises(ValueError):
            case_coherence.accept_finding(self.connection, self.case_id, self.profile, self.report().pending[0].key, ' ')
        self.assertTrue(self.report().pending)
