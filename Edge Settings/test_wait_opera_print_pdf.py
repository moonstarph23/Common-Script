import importlib.util
import os
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPT_PATH = Path(__file__).with_name("wait_opera_print_pdf.py")
SPEC = importlib.util.spec_from_file_location("wait_opera_print_pdf", SCRIPT_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class Clock:
    def __init__(self, now=100.0):
        self.now = now

    def time(self):
        return self.now

    def sleep(self, seconds):
        self.now += seconds


class WaitForOperaPrintPdfTests(unittest.TestCase):
    def run_wait(self, created_time, timeout=3):
        clock = Clock()
        candidate = os.path.join("cache", "OperaPrint.pdf")

        with patch.dict(MODULE.os.environ, {"LOCALAPPDATA": "local"}), \
                patch.object(MODULE.os.path, "isdir", return_value=True), \
                patch.object(MODULE.glob, "glob", side_effect=lambda pattern, recursive: [candidate] if "OperaPrint" in pattern else []), \
                patch.object(MODULE.os.path, "getctime", side_effect=created_time), \
                patch.object(MODULE.os.path, "getsize", return_value=128), \
                patch.object(MODULE.os, "open", return_value=10) as open_mock, \
                patch.object(MODULE.os, "close"), \
                patch.object(MODULE.time, "time", side_effect=clock.time), \
                patch.object(MODULE.time, "sleep", side_effect=clock.sleep):
            result = MODULE.wait_for_opera_print_pdf(timeout)

        return result, open_mock

    def test_ignores_pdf_created_before_function_started(self):
        result, open_mock = self.run_wait(lambda path: 99.0)

        self.assertEqual((False, None), result)
        open_mock.assert_not_called()

    def test_accepts_new_pdf_after_size_is_stable(self):
        result, open_mock = self.run_wait(lambda path: 101.0)

        self.assertEqual((True, os.path.join("cache", "OperaPrint.pdf")), result)
        open_mock.assert_called_once()

    def test_skips_pdf_when_creation_time_cannot_be_read(self):
        result, open_mock = self.run_wait(OSError("timestamp unavailable"))

        self.assertEqual((False, None), result)
        open_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main()
