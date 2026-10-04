"""Aceptación sintética del aprendizaje de CUPS de la mejora 6."""

import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path
from unittest.mock import patch

CORE = Path(__file__).resolve().parents[1] / "core"
if str(CORE) not in sys.path:
    sys.path.insert(0, str(CORE))

import case_ingestion
import community_discovery
import expedient_service
import gestor_bd
from community_cups import lookup_supply_point, normalize_cups
from document_text_service import TextExtraction


class CommunityCupsTest(unittest.TestCase):
    CUPS = "ES1234567890123456AA"

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.root = Path(self.directory.name)
        path = self.root / "gestion.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(path))
        self.connection = gestor_bd.conectar(str(path))
        self.communities = {}
        for code in ("701", "702"):
            community_id = gestor_bd.obtener_o_crear_comunidad(
                self.connection, code, f"Comunidad {code}",
            )
            case = expedient_service.create_case(
                self.connection, community_id, name="Enero 2026",
                start_date=date(2026, 1, 1), end_date=date(2026, 1, 31),
            )
            self.communities[code] = (community_id, case.id_case)

    def tearDown(self):
        self.connection.close()
        self.directory.cleanup()

    def add_invoice(self, code, filename, cups=None, supply="ELECTRICIDAD"):
        path = self.root / filename
        path.write_bytes(f"%PDF-1.4 factura sintetica {filename}".encode())
        values = {
            "tipo_suministro": supply, "fecha_inicio": "2026-01-01",
            "fecha_fin": "2026-01-31", "importe_total": "42.00",
            "num_factura": filename,
        }
        if cups:
            values["cups"] = cups
        case_id = self.communities[code][1]
        document = case_ingestion.add_document_to_case(
            self.connection, case_id, source_path=path,
            archive_root=self.root / "archive", document_kind="invoice",
            candidates=values, required_fields=(),
        ).document
        return case_id, document.id_document

    def confirm(self, code, filename, cups=None, supply="ELECTRICIDAD"):
        case_id, document_id = self.add_invoice(code, filename, cups, supply)
        case_ingestion.confirm_source_candidates(
            self.connection, case_id, document_id, confirmed_by="Operador",
        )
        return document_id

    def test_confirmation_teaches_cups_and_routes_later_pdf_without_code(self):
        first_id = self.confirm("701", "primera.pdf", f"ES 12345678 90123456 AA")
        point = lookup_supply_point(self.connection, self.CUPS)
        self.assertEqual(("701", "ELECTRICIDAD", "primera.pdf"),
                         (point.community_code, point.supply_type, point.source_name))
        self.assertEqual(self.CUPS, normalize_cups("ES 12345678 90123456 AA"))
        self.assertEqual(1, self.connection.execute(
            "SELECT COUNT(*) FROM learned_supply_points WHERE id_document=?", (first_id,),
        ).fetchone()[0])

        later = self.root / "factura_sin_codigo.pdf"
        later.write_bytes(b"%PDF-1.4 segunda factura sintetica")
        extraction = TextExtraction(
            f"Factura de luz. CUPS: {self.CUPS}", "pdf_text", (1,), {}, 1, False,
        )
        with patch.object(community_discovery, "get_document_text", return_value=extraction):
            proposal = community_discovery.build_global_intake(
                (later,), connection=self.connection,
            )
            discovered = community_discovery.discover_communities(
                (later,), connection=self.connection,
            )
            accepted, foreign = community_discovery.partition_sources_for_community(
                (later,), "702", connection=self.connection,
            )
        self.assertEqual(("701",), tuple(group.community_code for group in proposal.groups))
        self.assertEqual(("ELECTRICIDAD",), proposal.groups[0].supply_hints)
        self.assertEqual(("701",), tuple(item.code for item in discovered))
        self.assertEqual((), accepted)
        self.assertEqual((later,), foreign)

    def test_conflicting_confirmation_keeps_original_link_and_shows_both_sources(self):
        self.confirm("701", "origen.pdf", self.CUPS)
        case_id, document_id = self.add_invoice("702", "contradiccion.pdf", self.CUPS)
        with self.assertRaisesRegex(ValueError, "origen.pdf.*contradiccion.pdf"):
            case_ingestion.confirm_source_candidates(
                self.connection, case_id, document_id, confirmed_by="Operador",
            )
        self.assertEqual("701", lookup_supply_point(self.connection, self.CUPS).community_code)
        self.assertEqual(0, self.connection.execute(
            "SELECT COUNT(*) FROM facturas WHERE id_comunidad=?",
            (self.communities["702"][0],),
        ).fetchone()[0])

    def test_filename_that_disagrees_with_learned_cups_remains_unassigned(self):
        self.confirm("701", "origen.pdf", self.CUPS)
        conflict = self.root / "702_factura.pdf"
        conflict.write_bytes(b"%PDF-1.4 factura sintetica")
        extraction = TextExtraction(self.CUPS, "pdf_text", (1,), {}, 1, False)
        with patch.object(community_discovery, "get_document_text", return_value=extraction):
            proposal = community_discovery.build_global_intake(
                (conflict,), connection=self.connection,
            )
            accepted, foreign = community_discovery.partition_sources_for_community(
                (conflict,), "702", connection=self.connection,
            )
        self.assertFalse(proposal.groups)
        self.assertEqual((conflict,), proposal.unassigned_paths)
        self.assertEqual((), accepted)
        self.assertEqual((conflict,), foreign)

    def test_reconfirming_corrected_invoice_retires_old_cups(self):
        document_id = self.confirm("701", "corregida.pdf", self.CUPS)
        corrected = "ES9999999999999999BB"
        self.connection.execute("""UPDATE extraction_candidates
            SET value=?,validation_status='candidate'
            WHERE id_document=? AND field_name='cups'""", (corrected, document_id))
        case_ingestion.confirm_source_candidates(
            self.connection, self.communities["701"][1], document_id,
            confirmed_by="Operador",
        )
        self.assertIsNone(lookup_supply_point(self.connection, self.CUPS))
        self.assertEqual("701", lookup_supply_point(self.connection, corrected).community_code)
        self.connection.execute("""UPDATE extraction_candidates
            SET value='REFERENCIA',validation_status='candidate'
            WHERE id_document=? AND field_name='cups'""", (document_id,))
        case_ingestion.confirm_source_candidates(
            self.connection, self.communities["701"][1], document_id,
            confirmed_by="Operador",
        )
        self.assertIsNone(lookup_supply_point(self.connection, corrected))

    def test_unconfirmed_or_invalid_cups_never_teaches(self):
        self.add_invoice("701", "pendiente.pdf", self.CUPS)
        self.confirm("701", "referencia.pdf", "REFERENCIA 123")
        self.assertIsNone(lookup_supply_point(self.connection, self.CUPS))
        self.assertEqual(0, self.connection.execute(
            "SELECT COUNT(*) FROM learned_supply_points",
        ).fetchone()[0])


if __name__ == "__main__":
    unittest.main()
