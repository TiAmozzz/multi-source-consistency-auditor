from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "multi-source-consistency-auditor" / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from multi_source_consistency_auditor.cli import _validate_output_location
from multi_source_consistency_auditor.normalize import (
    normalize_amount,
    normalize_date,
    values_equal,
)
from multi_source_consistency_auditor.profile import ProfileError, load_profile
from multi_source_consistency_auditor.reports import write_reports
from multi_source_consistency_auditor.rules import AuditEngine


def file_hashes(root: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    for path in sorted(item for item in root.rglob("*") if item.is_file()):
        result[str(path)] = hashlib.sha256(path.read_bytes()).hexdigest()
    return result


class NormalizationTests(unittest.TestCase):
    def test_amount_normalization_preserves_integer_trailing_zeros(self) -> None:
        self.assertEqual("12000", normalize_amount("12,000.00"))
        self.assertEqual("12000", normalize_amount("1.2万元"))
        self.assertEqual("3500", normalize_amount("人民币 3,500 元"))
        self.assertTrue(values_equal("amount", "3500", "3500.01", tolerance=__import__("decimal").Decimal("0.01")))

    def test_date_normalization(self) -> None:
        self.assertEqual("2026-09-08", normalize_date("2026年9月8日"))
        self.assertEqual("2026-09-08", normalize_date("2026/09/08"))
        self.assertEqual("2026-09-08", normalize_date("20260908"))


class EndToEndTests(unittest.TestCase):
    def test_procurement_sample(self) -> None:
        profile_path = ROOT / "examples" / "procurement" / "profile.json"
        input_root = ROOT / "examples" / "procurement" / "input"
        before = file_hashes(input_root)
        result = AuditEngine(load_profile(profile_path)).run()
        after = file_hashes(input_root)

        self.assertEqual(before, after)
        self.assertTrue(result.to_dict()["integrity"]["source_files_unchanged"])
        self.assertEqual(5, len(result.entity_summaries))
        entity_by_id = {item["entity_id"]: item for item in result.entity_summaries}
        self.assertEqual("通过", entity_by_id["P-2026-001"]["status"])
        self.assertEqual("未通过", entity_by_id["P-2026-002"]["status"])

        amount_conflicts = [
            item
            for item in result.issue_list
            if item.code == "field_conflict"
            and item.entity_id == "P-2026-002"
            and item.field_name == "amount"
        ]
        self.assertEqual(1, len(amount_conflicts))
        sources = {item.source for item in amount_conflicts[0].evidence}
        self.assertEqual({"采购台账", "采购申请", "合同文件"}, sources)

        for entity_id in ("P-2026-003", "P-2026-004", "P-2026-005"):
            self.assertIn(
                "source_missing_record",
                {item.code for item in result.issue_list if item.entity_id == entity_id},
            )

    def test_reports_are_generated(self) -> None:
        result = AuditEngine(
            load_profile(ROOT / "examples" / "procurement" / "profile.json")
        ).run()
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = write_reports(
                result,
                temp_dir,
                ["html", "xlsx", "docx", "pdf", "markdown", "json", "csv"],
            )
            self.assertEqual(7, len(paths))
            for path in paths.values():
                self.assertTrue(path.exists(), path)
                self.assertGreater(path.stat().st_size, 0)
                self.assertEqual("differences-summary" + path.suffix.lower(), path.name)
            payload = json.loads(paths["json"].read_text(encoding="utf-8"))
            self.assertEqual("multi-source-consistency-auditor", payload["generator"]["name"])
            self.assertTrue(payload["integrity"]["source_files_unchanged"])
            word_text = paths["docx"].read_bytes()
            self.assertTrue(word_text.startswith(b"PK"))
            self.assertTrue(paths["pdf"].read_bytes().startswith(b"%PDF"))
            pdf_text = "\n".join(
                page.extract_text() or "" for page in PdfReader(str(paths["pdf"])).pages
            )
            for heading in ("总体汇总", "实体对照", "差异明细", "证据索引", "来源清单"):
                self.assertIn(heading, pdf_text)
            self.assertIn("原始资料只读", pdf_text)


class ValidationTests(unittest.TestCase):
    def test_duplicate_identifier_is_reported(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "records.csv").write_text(
                "编号,名称\nA-001,甲\nA-001,乙\n", encoding="utf-8-sig"
            )
            profile_data = {
                "name": "重复编号测试",
                "entity": "记录",
                "identifier": {"field": "id", "aliases": ["编号"]},
                "sources": [
                    {
                        "name": "源表",
                        "type": "table",
                        "paths": ["records.csv"],
                        "expected": False,
                    }
                ],
                "fields": {
                    "title": {"label": "名称", "aliases": ["名称"], "type": "text"}
                },
                "auto_consistent_fields": True,
            }
            profile_path = root / "profile.json"
            profile_path.write_text(
                json.dumps(profile_data, ensure_ascii=False), encoding="utf-8"
            )
            result = AuditEngine(load_profile(profile_path)).run()
            self.assertIn("duplicate_identifier", {item.code for item in result.issue_list})

    def test_multiple_excel_sheets_are_parsed(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            workbook = Workbook()
            first = workbook.active
            first.title = "第一批"
            first.append(["编号", "名称"])
            first.append(["A-001", "甲"])
            second = workbook.create_sheet("第二批")
            second.append(["编号", "名称"])
            second.append(["A-002", "乙"])
            workbook.save(root / "多表.xlsx")
            profile_data = {
                "name": "SHEET 测试",
                "entity": "记录",
                "identifier": {"field": "id", "aliases": ["编号"]},
                "sources": [
                    {
                        "name": "多表来源",
                        "type": "table",
                        "paths": ["多表.xlsx"],
                        "expected": True,
                    }
                ],
                "fields": {
                    "title": {"label": "名称", "aliases": ["名称"], "type": "text"}
                },
            }
            profile_path = root / "profile.json"
            profile_path.write_text(
                json.dumps(profile_data, ensure_ascii=False), encoding="utf-8"
            )
            result = AuditEngine(load_profile(profile_path)).run()
            self.assertEqual(
                ["A-001", "A-002"],
                sorted(item["entity_id"] for item in result.entity_summaries),
            )
            self.assertTrue(
                all("第" in record.location or "批" in record.location for record in result.records)
            )

    def test_invalid_profile_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "bad.json"
            path.write_text('{"name":"坏配置"}', encoding="utf-8")
            with self.assertRaises(ProfileError):
                load_profile(path)

    def test_output_inside_source_directory_is_rejected(self) -> None:
        profile = load_profile(ROOT / "examples" / "procurement" / "profile.json")
        with self.assertRaises(ValueError):
            _validate_output_location(
                profile,
                ROOT / "examples" / "procurement" / "input" / "合同文件" / "reports",
            )

    def test_output_directory_with_source_file_is_rejected(self) -> None:
        profile = load_profile(ROOT / "examples" / "procurement" / "profile.json")
        with self.assertRaises(ValueError):
            _validate_output_location(
                profile,
                ROOT / "examples" / "procurement" / "input",
            )


if __name__ == "__main__":
    unittest.main()
