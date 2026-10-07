"""Mejora 5: la bandeja global reparte por contenido y conserva la evidencia."""

import shutil
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
for folder in (PROJECT_ROOT / "core", PROJECT_ROOT / "tests"):
    if str(folder) not in sys.path:
        sys.path.insert(0, str(folder))

import case_ingestion
import community_discovery
import document_review
import expedient_service
import gestor_bd
import intake_routing
from test_source_preview import write_pdf
from source_analysis import SourceAnalysis

CIF_658, CIF_701 = "H10000008", "H10000016"


class IntakeByContentTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name).resolve()
        self.database = self.root / "gestion.db"
        with redirect_stdout(StringIO()):
            gestor_bd.crear_bd(str(self.database))
        self.connection = gestor_bd.conectar(str(self.database))
        self.addCleanup(self.connection.close)
        self.cases = {}
        for code, cif in (("658", CIF_658), ("701", CIF_701), ("702", None)):
            community = gestor_bd.obtener_o_crear_comunidad(self.connection, code, f"CP {code}", cif=cif)
            case = expedient_service.create_case(
                self.connection, community, name="Año 2026",
                start_date=date(2026, 1, 1), end_date=date(2026, 12, 31))
            self.cases[code] = (community, case.id_case)
        self.connection.commit()
        self.inbox = self.root / "correo"
        self.inbox.mkdir()

    def pdf(self, name, *lines):
        path = self.inbox / name
        write_pdf(path, [lines])
        return path

    def test_mixed_folder_is_split_by_cif_name_and_archived_file(self):
        by_cif = self.pdf("factura_luz.pdf", "Factura electricidad", f"Cliente: Comunidad de propietarios CIF {CIF_658}")
        other = self.pdf("factura_agua.pdf", "Factura agua", f"Titular CIF {CIF_701}")
        named = self.pdf("702_limpieza.pdf", "Factura limpieza sin CIF de comunidad")
        conflict = self.pdf("701_gas.pdf", "Factura gas", f"Cliente CIF {CIF_658}")
        unknown = self.pdf("presupuesto.pdf", "Presupuesto sin datos de comunidad")
        # El mismo archivo ya está en un expediente de 702: se reconoce aunque se renombre.
        archived = self.root / "original.pdf"
        write_pdf(archived, [["Factura ascensor", "Referencia 99"]])
        case_ingestion.add_document_to_case(
            self.connection, self.cases["702"][1], source_path=archived,
            archive_root=self.root / "archivo", document_kind="invoice", candidates={}, required_fields=())
        renamed = self.inbox / "copia_reenviada.pdf"
        shutil.copy2(archived, renamed)

        proposal = community_discovery.build_global_intake(
            sorted(self.inbox.iterdir()), connection=self.connection)

        groups = {group.community_code: group for group in proposal.groups}
        self.assertEqual({"658", "701", "702"}, set(groups))
        self.assertEqual((by_cif,), groups["658"].source_paths)
        self.assertEqual((other,), groups["701"].source_paths)
        self.assertEqual({named, renamed}, set(groups["702"].source_paths))
        evidence = dict(groups["658"].evidence)
        self.assertIn(f"CIF: {CIF_658} de la comunidad 658", evidence[by_cif])
        self.assertIn("archivo ya archivado", " ".join(dict(groups["702"].evidence)[renamed]))
        self.assertIn("1 por nombre", groups["702"].evidence_summary)

        self.assertEqual({conflict, unknown}, set(proposal.unassigned_paths))
        self.assertIn("contradictorias", proposal.reason_for(conflict))
        self.assertIn("701 por nombre", proposal.reason_for(conflict))
        self.assertIn("658 por CIF", proposal.reason_for(conflict))
        self.assertIn("Sin código", proposal.reason_for(unknown))

    def test_adding_to_a_case_sets_aside_files_of_another_community_by_content(self):
        by_cif = self.pdf("factura_luz.pdf", f"Cliente CIF {CIF_658}")
        conflict = self.pdf("701_gas.pdf", f"Cliente CIF {CIF_658}")
        plain = self.pdf("anexo.pdf", "Anexo sin datos")
        accepted, foreign = community_discovery.partition_sources_for_community(
            (by_cif, conflict, plain), "701", connection=self.connection)
        self.assertEqual((plain,), accepted)
        self.assertEqual({by_cif, conflict}, set(foreign))

    def test_groups_and_manual_choices_reach_each_open_case(self):
        by_cif = self.pdf("factura_luz.pdf", "Factura electricidad", f"Cliente CIF {CIF_658}")
        unknown = self.pdf("presupuesto.pdf", "Presupuesto sin datos de comunidad")
        proposal = community_discovery.build_global_intake((by_cif, unknown), connection=self.connection)
        self.assertEqual((unknown,), proposal.unassigned_paths)

        assignments = intake_routing.assignments_from_proposal(proposal, {unknown: "702"})
        targets = {target.code: target for target in intake_routing.plan_routes(self.connection, assignments)}
        self.assertEqual(self.cases["658"][1], targets["658"].case_id)
        self.assertIn("Año 2026", targets["658"].case_label)
        self.assertIsNone(targets["658"].blocked_reason)

        results = [intake_routing.ingest_route(self.database, target, archive_root=self.root / "archivo")
                   for target in targets.values()]
        self.assertEqual([1, 1], [result.created for result in results])
        self.assertEqual([[], []], [result.errors for result in results])
        names = {row[0]: row[1] for row in self.connection.execute(
            "SELECT original_name, id_case FROM source_documents")}
        self.assertEqual(self.cases["658"][1], names["factura_luz.pdf"])
        self.assertEqual(self.cases["702"][1], names["presupuesto.pdf"])

        # Repetir el reparto no duplica documentos.
        again = intake_routing.ingest_route(self.database, targets["658"], archive_root=self.root / "archivo")
        self.assertEqual((0, 1), (again.created, again.duplicates))

    def test_route_without_open_case_or_community_is_blocked(self):
        self.connection.execute("UPDATE regularization_cases SET estado='closed' WHERE id_case=?",
                                (self.cases["701"][1],))
        self.connection.commit()
        targets = {t.code: t for t in intake_routing.plan_routes(
            self.connection, {"701": [self.inbox / "x.pdf"], "999": [self.inbox / "y.pdf"]})}
        self.assertIn("expediente abierto", targets["701"].blocked_reason)
        self.assertIn("no está dada de alta", targets["999"].blocked_reason)
        result = intake_routing.ingest_route(self.database, targets["999"], archive_root=self.root / "archivo")
        self.assertEqual(0, result.created)
        self.assertTrue(result.errors)

    def test_manual_conflict_is_audited_and_never_applied_without_review(self):
        conflict = self.pdf("701_gas.pdf", "Factura gas", f"Cliente CIF {CIF_658}")
        proposal = community_discovery.build_global_intake((conflict,), connection=self.connection)
        reason = proposal.reason_for(conflict)
        self.assertIn("contradictorias", reason)
        assignments = intake_routing.assignments_from_proposal(proposal, {conflict: "658"})
        decision = intake_routing.ManualAssignment(conflict, "658", reason, "CIF y titular comprobados")
        target = intake_routing.plan_routes(
            self.connection, assignments, manual_assignments=(decision,),
        )[0]
        complete = SourceAnalysis.invoice({
            "tipo_suministro": "GAS", "fecha_inicio": "2026-01-01",
            "fecha_fin": "2026-01-31", "importe_total": "120.00",
        })
        analyse = lambda _path, **_kwargs: complete
        result = intake_routing.ingest_route(
            self.database, target, archive_root=self.root / "archivo", analyser=analyse,
        )
        self.assertEqual((1, (), []), (result.created, result.foreign, result.errors))
        row = self.connection.execute(
            "SELECT id_document,status FROM source_documents WHERE original_name=?", (conflict.name,),
        ).fetchone()
        self.assertNotEqual("validated", row["status"])
        issue = self.connection.execute(
            "SELECT id_issue,message,detected_value FROM review_issues WHERE id_document=? AND code=?",
            (row["id_document"], "manual_community_assignment"),
        ).fetchone()
        self.assertIn(reason, issue["message"])
        self.assertIn("CIF y titular comprobados", issue["message"])
        with self.assertRaisesRegex(ValueError, "comunidad"):
            document_review.resolve_issue(
                self.connection, issue["id_issue"], value="701", reason="Equivocación",
            )
        again = intake_routing.ingest_route(
            self.database, target, archive_root=self.root / "archivo", analyser=analyse,
        )
        self.assertEqual((0, 1), (again.created, again.duplicates))
        self.assertNotEqual("validated", self.connection.execute(
            "SELECT status FROM source_documents WHERE id_document=?", (row["id_document"],)
        ).fetchone()[0])
        document_review.resolve_issue(
            self.connection, issue["id_issue"], value="658", reason="CIF contrastado en el original",
        )
        third = intake_routing.ingest_route(
            self.database, target, archive_root=self.root / "archivo", analyser=analyse,
        )
        self.assertEqual((0, 1), (third.created, third.duplicates))
        self.assertEqual(0, self.connection.execute(
            """SELECT COUNT(*) FROM review_issues WHERE id_document=?
               AND code='manual_community_assignment' AND status='open'""",
            (row["id_document"],),
        ).fetchone()[0])

    def test_manual_conflict_needs_reason_and_fresh_evidence(self):
        conflict = self.pdf("701_gas.pdf", f"Cliente CIF {CIF_658}")
        proposal = community_discovery.build_global_intake((conflict,), connection=self.connection)
        reason = proposal.reason_for(conflict)
        assignments = {"658": [conflict]}
        missing_reason = intake_routing.ManualAssignment(conflict, "658", reason)
        target = intake_routing.plan_routes(
            self.connection, assignments, manual_assignments=(missing_reason,),
        )[0]
        rejected = intake_routing.ingest_route(
            self.database, target, archive_root=self.root / "archivo",
        )
        self.assertEqual(0, rejected.created)
        self.assertTrue(rejected.errors)
        stale = intake_routing.ManualAssignment(conflict, "658", "Otra evidencia", "Revisado")
        target = intake_routing.plan_routes(
            self.connection, assignments, manual_assignments=(stale,),
        )[0]
        rejected = intake_routing.ingest_route(
            self.database, target, archive_root=self.root / "archivo",
        )
        self.assertEqual(0, rejected.created)
        self.assertTrue(rejected.errors)


if __name__ == "__main__":
    unittest.main()
