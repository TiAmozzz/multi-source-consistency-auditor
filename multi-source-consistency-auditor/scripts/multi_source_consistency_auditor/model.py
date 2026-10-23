"""Shared data structures used by parsers, rules, and reports."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class Evidence:
    source: str
    file: str
    location: str
    excerpt: str

    def to_dict(self) -> dict[str, str]:
        return {
            "source": self.source,
            "file": self.file,
            "location": self.location,
            "excerpt": self.excerpt,
        }


@dataclass(frozen=True)
class Fact:
    field: str
    raw_value: str
    normalized_value: str
    evidence: Evidence

    def to_dict(self) -> dict[str, Any]:
        return {
            "field": self.field,
            "raw_value": self.raw_value,
            "normalized_value": self.normalized_value,
            "evidence": self.evidence.to_dict(),
        }


@dataclass
class Record:
    record_id: str
    source: str
    file: str
    location: str
    facts: dict[str, list[Fact]] = field(default_factory=dict)
    metadata: dict[str, Any] = field(default_factory=dict)

    def add_fact(self, fact: Fact) -> None:
        self.facts.setdefault(fact.field, []).append(fact)

    def facts_for(self, field_name: str) -> list[Fact]:
        return self.facts.get(field_name, [])

    def to_dict(self) -> dict[str, Any]:
        return {
            "record_id": self.record_id,
            "source": self.source,
            "file": self.file,
            "location": self.location,
            "metadata": self.metadata,
            "facts": {
                key: [fact.to_dict() for fact in values]
                for key, values in sorted(self.facts.items())
            },
        }


@dataclass
class Issue:
    severity: str
    code: str
    entity_id: str
    message: str
    field_name: str | None = None
    expected: str = ""
    actual: str = ""
    evidence: list[Evidence] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "severity": self.severity,
            "code": self.code,
            "entity_id": self.entity_id,
            "field": self.field_name or "",
            "message": self.message,
            "expected": self.expected,
            "actual": self.actual,
            "evidence": [item.to_dict() for item in self.evidence],
        }


@dataclass
class SourceInventory:
    source: str
    file: str
    source_type: str
    status: str
    records: int = 0

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "file": self.file,
            "type": self.source_type,
            "status": self.status,
            "records": self.records,
        }


@dataclass
class AuditResult:
    profile_name: str
    entity_name: str
    issue_list: list[Issue]
    records: list[Record]
    source_inventory: list[SourceInventory]
    entity_summaries: list[dict[str, Any]]
    started_at: str
    finished_at: str
    input_files_checked: int = 0
    modified_input_files: list[str] = field(default_factory=list)

    @property
    def error_count(self) -> int:
        return sum(item.severity == "error" for item in self.issue_list)

    @property
    def warning_count(self) -> int:
        return sum(item.severity == "warning" for item in self.issue_list)

    @property
    def info_count(self) -> int:
        return sum(item.severity == "info" for item in self.issue_list)

    @property
    def status(self) -> str:
        if self.error_count:
            return "未通过"
        if self.warning_count:
            return "需复核"
        return "通过"

    def to_dict(self) -> dict[str, Any]:
        return {
            "profile": self.profile_name,
            "entity": self.entity_name,
            "status": self.status,
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "integrity": {
                "source_files_checked": self.input_files_checked,
                "source_files_modified": self.modified_input_files,
                "source_files_unchanged": not self.modified_input_files,
                "read_only": True,
                "parser_input": "temporary_copy",
                "network_used": False,
            },
            "summary": {
                "sources": len({item.source for item in self.source_inventory}),
                "files": len(self.source_inventory),
                "records": len(self.records),
                "entities": len(self.entity_summaries),
                "passed_entities": sum(
                    item["status"] == "通过" for item in self.entity_summaries
                ),
                "review_entities": sum(
                    item["status"] != "通过" for item in self.entity_summaries
                ),
                "errors": self.error_count,
                "warnings": self.warning_count,
                "infos": self.info_count,
            },
            "sources": [item.to_dict() for item in self.source_inventory],
            "entities": self.entity_summaries,
            "issues": [item.to_dict() for item in self.issue_list],
            "records": [item.to_dict() for item in self.records],
        }
