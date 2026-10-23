from __future__ import annotations

import json
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "multi-source-consistency-auditor" / "SKILL.md"


class SkillStructureTests(unittest.TestCase):
    def test_frontmatter_and_description(self) -> None:
        text = SKILL.read_text(encoding="utf-8")
        match = re.match(r"^---\r?\n(.*?)\r?\n---\r?\n", text, re.DOTALL)
        self.assertIsNotNone(match)
        frontmatter = match.group(1)
        keys = re.findall(r"^([a-z][a-z0-9_-]*):", frontmatter, re.MULTILINE)
        self.assertEqual(["name", "description"], keys)

        name_match = re.search(r'^name:\s*"?([a-z0-9-]+)"?\s*$', frontmatter, re.MULTILINE)
        description_match = re.search(
            r'^description:\s*("(?:[^"\\]|\\.)*")\s*$',
            frontmatter,
            re.MULTILINE,
        )
        self.assertIsNotNone(name_match)
        self.assertIsNotNone(description_match)
        name = name_match.group(1)
        description = json.loads(description_match.group(1))

        self.assertEqual(SKILL.parent.name, name)
        self.assertRegex(name, r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
        self.assertLessEqual(len(name), 64)
        self.assertLessEqual(len(description), 1024)
        self.assertGreaterEqual(description.count("；"), 2)
        triggers = description.split("触发词：", 1)[1]
        self.assertGreaterEqual(len([part for part in triggers.split("、") if part]), 6)

    def test_read_only_boundary_is_explicit(self) -> None:
        text = SKILL.read_text(encoding="utf-8")
        for phrase in ("原始资料只读", "不得联网", "临时副本", "SHA-256"):
            self.assertIn(phrase, text)


if __name__ == "__main__":
    unittest.main()
