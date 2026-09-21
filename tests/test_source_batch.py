import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CORE_DIR = PROJECT_ROOT / "core"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from source_analysis import SourceAnalysis
from source_batch import analyse_batch


class SourceBatchTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.directory = Path(self.temp.name)
        self.database_path = self.directory / "gestion.db"

    def tearDown(self):
        self.temp.cleanup()

    def _file(self, name, content=None):
        path = self.directory / name
        path.write_bytes(content or name.encode("utf-8"))
        return path

    def test_batch_continues_after_one_document_error(self):
        events = []

        def analyser(path, **_kwargs):
            if path.name == "broken.pdf":
                raise ValueError("PDF ilegible")
            return SourceAnalysis.invoice({
                "tipo_suministro": "GAS",
                "fecha_inicio": "2026-01-01",
                "fecha_fin": "2026-01-31",
                "importe_total": "10.00",
            })

        results = analyse_batch(
            (self._file("ok.pdf"), self._file("broken.pdf"), self._file("ok2.pdf")),
            community_code="658",
            database_path=self.database_path,
            analyser=analyser,
            max_workers=2,
            progress=events.append,
        )
        self.assertEqual(3, len(results))
        self.assertEqual(1, sum(item.error is not None for item in results))
        self.assertEqual("completed", events[-1].phase)

    def test_batch_never_uses_more_than_three_workers(self):
        with mock.patch("source_batch.ThreadPoolExecutor") as executor:
            analyse_batch(
                (), community_code="658", database_path=self.database_path
            )
        executor.assert_called_once_with(
            max_workers=3, thread_name_prefix="source-analysis"
        )

    def test_duplicate_content_is_analysed_once_but_keeps_input_order(self):
        first = self._file("first.pdf", b"same")
        second = self._file("second.pdf", b"same")
        calls = []

        def analyser(path, **_kwargs):
            calls.append(path)
            return SourceAnalysis.invoice({"importe_total": "10.00"})

        results = analyse_batch(
            (first, second),
            community_code="658",
            database_path=self.database_path,
            analyser=analyser,
        )

        self.assertEqual(1, len(calls))
        self.assertEqual((first, second), tuple(item.path for item in results))


if __name__ == "__main__":
    unittest.main()
