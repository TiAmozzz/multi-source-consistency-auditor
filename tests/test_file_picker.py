from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools.file_picker import clean_filename, detect_type


class FilePickerTests(unittest.TestCase):
    def test_filename_is_sanitized(self) -> None:
        self.assertEqual("001_a_b.csv", clean_filename("../a:b.csv", 1))

    def test_field_types_are_inferred(self) -> None:
        self.assertEqual("amount", detect_type("合同金额", "12000"))
        self.assertEqual("date", detect_type("签订日期", "2026-01-01"))
        self.assertEqual("number", detect_type("采购数量", "3"))
        self.assertEqual("text", detect_type("供应商", "甲公司"))


if __name__ == "__main__":
    unittest.main()
