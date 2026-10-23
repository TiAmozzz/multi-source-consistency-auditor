"""Deterministic cross-source consistency rules."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable

from .model import AuditResult, Fact, Issue, Record
from .normalize import display_number, values_equal
from .parsers import collect_records, snapshot_source_files
from .profile import Profile, ProfileError


class AuditEngine:
    """Run the profile's checks over all source records."""

    def __init__(self, profile: Profile):
        self.profile = profile
        self._validate_checks()

    def _validate_checks(self) -> None:
        source_names = {item.name for item in self.profile.sources}
        for index, check in enumerate(self.profile.checks):
            check_type = check["type"]
            severity = check.get("severity", "warning")
            if severity not in {"error", "warning", "info"}:
                raise ProfileError(f"checks[{index}].severity 必须是 error、warning 或 info")
            configured_sources = check.get("sources", [])
            if not isinstance(configured_sources, list) or not all(
                isinstance(item, str) for item in configured_sources
            ):
                raise ProfileError(f"checks[{index}].sources 必须是字符串数组")
            missing_sources = set(configured_sources) - source_names
            if missing_sources:
                raise ProfileError(
                    f"checks[{index}] 引用了不存在的来源：{', '.join(sorted(missing_sources))}"
                )
            field_keys = (
                [check.get("field")]
                if check_type in {"consistent", "required"}
                else [check.get("start"), check.get("end")]
            )
            for field_name in field_keys:
                if field_name not in self.profile.fields:
                    raise ProfileError(f"checks[{index}] 引用了不存在的字段：{field_name}")

    def run(self) -> AuditResult:
        started_at = datetime.now().astimezone().isoformat(timespec="seconds")
        before_hashes = snapshot_source_files(self.profile)
        records, inventory, issues = collect_records(self.profile)
        after_hashes = snapshot_source_files(self.profile)
        modified_files = sorted(
            str(path)
            for path in set(before_hashes) | set(after_hashes)
            if before_hashes.get(path) != after_hashes.get(path)
        )
        if modified_files:
            issues.append(
                Issue(
                    severity="error",
                    code="source_file_modified",
                    entity_id="来源",
                    message="检测到原始资料在核验期间被外部改动",
                    actual="；".join(modified_files),
                )
            )
        if not records:
            issues.append(
                Issue(
                    severity="error",
                    code="no_records",
                    entity_id="全部",
                    message="未从任何来源解析出可用记录",
                )
            )

        groups: dict[str, list[Record]] = defaultdict(list)
        for record in records:
            groups[record.record_id].append(record)

        expected_sources = {
            item.name for item in self.profile.sources if item.expected
        }
        for entity_id, entity_records in sorted(groups.items()):
            if entity_id.startswith("未识别ID::"):
                continue
            sources_present = {item.source for item in entity_records}
            for source_name in sorted(expected_sources - sources_present):
                issues.append(
                    Issue(
                        severity=self.profile.coverage_severity,
                        code="source_missing_record",
                        entity_id=entity_id,
                        message=f"来源“{source_name}”缺少该实体的记录",
                        expected=f"{source_name} 中存在同一 {self.profile.identifier_field}",
                        actual="未找到记录",
                    )
                )

        # Duplicate keys within one source are highlighted separately from cross-source conflicts.
        for entity_id, entity_records in sorted(groups.items()):
            by_source: dict[str, list[Record]] = defaultdict(list)
            for record in entity_records:
                by_source[record.source].append(record)
            for source_name, source_records in by_source.items():
                if len(source_records) > 1 and not entity_id.startswith("未识别ID::"):
                    issues.append(
                        Issue(
                            severity="error",
                            code="duplicate_identifier",
                            entity_id=entity_id,
                            message=(
                                f"来源“{source_name}”中同一编号出现 "
                                f"{len(source_records)} 次"
                            ),
                            actual=", ".join(
                                f"{item.file} / {item.location}" for item in source_records
                            ),
                        )
                    )

        for check in self.profile.checks:
            check_type = check["type"]
            if check_type == "consistent":
                self._check_consistency(groups, check, issues)
            elif check_type == "required":
                self._check_required(groups, check, issues)
            elif check_type == "date_order":
                self._check_date_order(groups, check, issues)

        if self.profile.auto_consistent_fields:
            explicit_fields = {
                item.get("field") for item in self.profile.checks if item["type"] == "consistent"
            }
            for field_name in self.profile.fields:
                if field_name not in explicit_fields:
                    self._check_consistency(
                        groups,
                        {
                            "type": "consistent",
                            "field": field_name,
                            "severity": "warning",
                        },
                        issues,
                    )

        entity_summaries = self._build_entity_summaries(groups, issues)
        finished_at = datetime.now().astimezone().isoformat(timespec="seconds")
        return AuditResult(
            profile_name=self.profile.name,
            entity_name=self.profile.entity_name,
            issue_list=issues,
            records=records,
            source_inventory=inventory,
            entity_summaries=entity_summaries,
            started_at=started_at,
            finished_at=finished_at,
            input_files_checked=len(after_hashes),
            modified_input_files=modified_files,
        )

    def _selected_facts(
        self,
        records: Iterable[Record],
        field_name: str,
        source_names: Iterable[str] | None = None,
    ) -> list[Fact]:
        allowed = set(source_names or [])
        facts: list[Fact] = []
        for record in records:
            if allowed and record.source not in allowed:
                continue
            facts.extend(record.facts_for(field_name))
        return facts

    def _cluster_facts(self, field_name: str, facts: list[Fact], tolerance: Decimal) -> list[list[Fact]]:
        field_type = self.profile.fields[field_name].field_type
        clusters: list[list[Fact]] = []
        for fact in facts:
            for cluster in clusters:
                if values_equal(
                    field_type,
                    cluster[0].normalized_value,
                    fact.normalized_value,
                    tolerance,
                ):
                    cluster.append(fact)
                    break
            else:
                clusters.append([fact])
        return clusters

    def _check_consistency(
        self,
        groups: dict[str, list[Record]],
        check: dict[str, Any],
        issues: list[Issue],
    ) -> None:
        field_name = check["field"]
        source_names = check.get("sources") or None
        severity = check.get("severity", "warning")
        tolerance_raw = check.get("tolerance", self.profile.tolerances.get(field_name, "0"))
        try:
            tolerance = Decimal(str(tolerance_raw))
        except InvalidOperation as exc:
            raise ProfileError(f"字段 {field_name} 的 tolerance 不是有效数字") from exc

        for entity_id, entity_records in sorted(groups.items()):
            if entity_id.startswith("未识别ID::"):
                continue
            facts = self._selected_facts(entity_records, field_name, source_names)
            if not facts:
                continue
            source_count = len({fact.evidence.source for fact in facts})
            if source_count < 2:
                continue
            clusters = self._cluster_facts(field_name, facts, tolerance)
            if len(clusters) <= 1:
                continue
            field_spec = self.profile.fields[field_name]
            display_values: list[str] = []
            evidence_by_source: dict[str, Fact] = {}
            for cluster in clusters:
                for fact in cluster:
                    label = (
                        f"{fact.evidence.source}={fact.raw_value}"
                        f"（标准化为 {display_number(fact.normalized_value)}）"
                    )
                    if label not in display_values:
                        display_values.append(label)
                    evidence_by_source.setdefault(fact.evidence.source, fact)
            issues.append(
                Issue(
                    severity=severity,
                    code="field_conflict",
                    entity_id=entity_id,
                    field_name=field_name,
                    message=f"{field_spec.label}在不同来源中不一致",
                    expected=display_values[0] if display_values else "",
                    actual="；".join(display_values[1:]),
                    evidence=[fact.evidence for fact in evidence_by_source.values()],
                )
            )

    def _check_required(
        self,
        groups: dict[str, list[Record]],
        check: dict[str, Any],
        issues: list[Issue],
    ) -> None:
        field_name = check["field"]
        source_names = check.get("sources") or [item.name for item in self.profile.sources]
        severity = check.get("severity", "warning")
        field_spec = self.profile.fields[field_name]
        for entity_id, entity_records in sorted(groups.items()):
            if entity_id.startswith("未识别ID::"):
                continue
            for source_name in source_names:
                source_records = [item for item in entity_records if item.source == source_name]
                if not source_records:
                    continue
                if any(record.facts_for(field_name) for record in source_records):
                    continue
                issues.append(
                    Issue(
                        severity=severity,
                        code="required_field_missing",
                        entity_id=entity_id,
                        field_name=field_name,
                        message=f"来源“{source_name}”缺少必要字段：{field_spec.label}",
                        expected=f"提供{field_spec.label}",
                        actual="空值",
                        evidence=[
                            item.evidence
                            for record in source_records
                            for fact_values in record.facts.values()
                            for item in fact_values[:1]
                        ][:3],
                    )
                )

    def _check_date_order(
        self,
        groups: dict[str, list[Record]],
        check: dict[str, Any],
        issues: list[Issue],
    ) -> None:
        start_field = check["start"]
        end_field = check["end"]
        source_names = check.get("sources") or None
        severity = check.get("severity", "warning")
        start_label = self.profile.fields[start_field].label
        end_label = self.profile.fields[end_field].label

        for entity_id, entity_records in sorted(groups.items()):
            if entity_id.startswith("未识别ID::"):
                continue
            starts = self._selected_facts(entity_records, start_field, source_names)
            ends = self._selected_facts(entity_records, end_field, source_names)
            if not starts or not ends:
                continue
            latest_start = max(starts, key=lambda item: item.normalized_value)
            earliest_end = min(ends, key=lambda item: item.normalized_value)
            if latest_start.normalized_value > earliest_end.normalized_value:
                issues.append(
                    Issue(
                        severity=severity,
                        code="date_order_violation",
                        entity_id=entity_id,
                        field_name=start_field,
                        message=f"{start_label}晚于{end_label}",
                        expected=f"{start_label}不晚于{end_label}",
                        actual=(
                            f"{start_label}={latest_start.raw_value}；"
                            f"{end_label}={earliest_end.raw_value}"
                        ),
                        evidence=[latest_start.evidence, earliest_end.evidence],
                    )
                )

    def _build_entity_summaries(
        self,
        groups: dict[str, list[Record]],
        issues: list[Issue],
    ) -> list[dict[str, Any]]:
        issue_counts: dict[str, list[Issue]] = defaultdict(list)
        for issue in issues:
            issue_counts[issue.entity_id].append(issue)

        summaries: list[dict[str, Any]] = []
        for entity_id, entity_records in sorted(groups.items()):
            entity_issues = issue_counts.get(entity_id, [])
            status = (
                "未通过"
                if any(item.severity == "error" for item in entity_issues)
                else "需复核"
                if entity_issues
                else "通过"
            )
            values: dict[str, list[dict[str, str]]] = {}
            for field_name, field_spec in self.profile.fields.items():
                seen: set[tuple[str, str]] = set()
                field_values: list[dict[str, str]] = []
                for record in entity_records:
                    for fact in record.facts_for(field_name):
                        key = (fact.evidence.source, fact.normalized_value)
                        if key in seen:
                            continue
                        seen.add(key)
                        field_values.append(
                            {
                                "value": fact.raw_value,
                                "normalized": fact.normalized_value,
                                "source": fact.evidence.source,
                            }
                        )
                if field_values:
                    values[field_name] = field_values
            summaries.append(
                {
                    "entity_id": entity_id,
                    "status": status,
                    "sources": sorted({item.source for item in entity_records}),
                    "files": sorted({item.file for item in entity_records}),
                    "fields": values,
                    "labels": {
                        key: value.label for key, value in self.profile.fields.items()
                    },
                    "issue_count": len(entity_issues),
                    "error_count": sum(item.severity == "error" for item in entity_issues),
                    "warning_count": sum(item.severity == "warning" for item in entity_issues),
                }
            )
        return summaries
