import json
import sys
import tempfile
import unittest
from dataclasses import asdict
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
for directory in (CORE_DIR, SCRIPTS_DIR):
    if str(directory) not in sys.path:
        sys.path.insert(0, str(directory))

from audit_provider_catalog import audit_documents
from provider_registry import provider_registry_from_payload


class ProviderCatalogAuditTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.directory = Path(self.temp.name)
        self.registry = provider_registry_from_payload({"proveedores": {
            key: {
                "tax_ids": [],
                "aliases": [key],
                "document_types": ["invoice"],
                "service_family": "MANTENIMIENTO",
                "extractor_family": "standard_spanish_invoice",
                "required_signatures": [r"\bFACTURA\b"],
                "excluded_signatures": [r"\bPRESUPUESTO\b"],
            }
            for key in ("TIERSAN", "ECHEMAN")
        }})

    def tearDown(self):
        self.temp.cleanup()

    def test_audit_report_contains_counts_but_no_document_text(self):
        path = self.directory / "tiersan.pdf"
        path.write_bytes(b"synthetic-pdf")
        report = audit_documents(
            (path,),
            registry=self.registry,
            text_loader=lambda _path: "TIERSAN FACTURA F-1 Total 121,00 €",
        )
        self.assertEqual(1, report.unique_documents)
        self.assertEqual(1, report.recognised_invoices)
        self.assertNotIn("TIERSAN FACTURA", json.dumps(asdict(report)))

    def test_duplicate_sha_is_counted_once(self):
        first = self.directory / "a.pdf"
        second = self.directory / "b.pdf"
        first.write_bytes(b"same-synthetic-pdf")
        second.write_bytes(b"same-synthetic-pdf")
        report = audit_documents(
            (first, second),
            registry=self.registry,
            text_loader=lambda _path: "ECHEMAN FACTURA F-1 Total 20,00 €",
        )
        self.assertEqual(1, report.unique_documents)
        self.assertEqual(2, report.input_paths)


if __name__ == "__main__":
    unittest.main()
