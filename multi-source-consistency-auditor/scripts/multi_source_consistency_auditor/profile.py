"""Load and validate audit profiles."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


class ProfileError(ValueError):
    """Raised for invalid or incomplete audit profiles."""


@dataclass(frozen=True)
class FieldSpec:
    name: str
    label: str
    aliases: tuple[str, ...]
    field_type: str


@dataclass(frozen=True)
class SourceSpec:
    name: str
    kind: str
    paths: tuple[str, ...]
    expected: bool = True
    include: tuple[str, ...] = ()


@dataclass
class Profile:
    name: str
    entity_name: str
    identifier_field: str
    identifier_aliases: tuple[str, ...]
    identifier_pattern: str | None
    sources: tuple[SourceSpec, ...]
    fields: dict[str, FieldSpec]
    checks: tuple[dict[str, Any], ...]
    coverage_severity: str
    auto_consistent_fields: bool
    tolerances: dict[str, str]
    base_dir: Path
    path: Path
    raw: dict[str, Any] = field(default_factory=dict)

    def resolve_paths(self, source: SourceSpec) -> list[Path]:
        resolved: list[Path] = []
        for item in source.paths:
            candidate = Path(item)
            if not candidate.is_absolute():
                candidate = self.base_dir / candidate
            matches = sorted(candidate.parent.glob(candidate.name)) if candidate.parent.exists() else []
            if candidate.exists():
                resolved.append(candidate)
            elif matches:
                resolved.extend(matches)
        unique: list[Path] = []
        seen: set[Path] = set()
        for item in resolved:
            absolute = item.resolve()
            if absolute not in seen:
                seen.add(absolute)
                unique.append(absolute)
        return unique


def _string_list(value: Any, location: str, allow_empty: bool = False) -> tuple[str, ...]:
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, list) or not all(isinstance(item, str) and item.strip() for item in value):
        raise ProfileError(f"{location} 必须是字符串或非空字符串数组")
    if not value and not allow_empty:
        raise ProfileError(f"{location} 不能为空")
    return tuple(item.strip() for item in value)


def _require_string(data: dict[str, Any], key: str) -> str:
    value = data.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ProfileError(f"缺少非空字符串配置项：{key}")
    return value.strip()


def load_profile(path: str | Path) -> Profile:
    profile_path = Path(path).resolve()
    if not profile_path.exists():
        raise ProfileError(f"核验配置不存在：{profile_path}")
    try:
        data = json.loads(profile_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProfileError(f"无法读取核验配置：{exc}") from exc
    if not isinstance(data, dict):
        raise ProfileError("核验配置顶层必须是 JSON 对象")

    name = _require_string(data, "name")
    entity_name = data.get("entity", "业务实体")
    if not isinstance(entity_name, str) or not entity_name.strip():
        raise ProfileError("entity 必须是非空字符串")

    identifier = data.get("identifier")
    if not isinstance(identifier, dict):
        raise ProfileError("identifier 必须包含 field、aliases，可选 pattern")
    identifier_field = _require_string(identifier, "field")
    identifier_aliases = _string_list(identifier.get("aliases"), "identifier.aliases")
    identifier_pattern = identifier.get("pattern")
    if identifier_pattern is not None:
        if not isinstance(identifier_pattern, str) or not identifier_pattern:
            raise ProfileError("identifier.pattern 必须是非空正则表达式")
        try:
            re.compile(identifier_pattern)
        except re.error as exc:
            raise ProfileError(f"identifier.pattern 不是有效正则表达式：{exc}") from exc

    raw_sources = data.get("sources")
    if not isinstance(raw_sources, list) or not raw_sources:
        raise ProfileError("sources 必须是非空数组")
    sources: list[SourceSpec] = []
    source_names: set[str] = set()
    for index, item in enumerate(raw_sources):
        if not isinstance(item, dict):
            raise ProfileError(f"sources[{index}] 必须是对象")
        source_name = _require_string(item, "name")
        if source_name in source_names:
            raise ProfileError(f"来源名称重复：{source_name}")
        source_names.add(source_name)
        kind = item.get("type", "auto")
        if kind not in {"auto", "table", "document"}:
            raise ProfileError(f"sources[{index}].type 必须是 auto、table 或 document")
        raw_paths = item.get("paths", item.get("path"))
        paths = _string_list(raw_paths, f"sources[{index}].paths")
        include = item.get("include", [])
        include = _string_list(include, f"sources[{index}].include", allow_empty=True)
        expected = item.get("expected", True)
        if not isinstance(expected, bool):
            raise ProfileError(f"sources[{index}].expected 必须是布尔值")
        sources.append(
            SourceSpec(
                name=source_name,
                kind=kind,
                paths=paths,
                expected=expected,
                include=include,
            )
        )

    raw_fields = data.get("fields")
    if not isinstance(raw_fields, dict) or not raw_fields:
        raise ProfileError("fields 必须是非空对象")
    fields: dict[str, FieldSpec] = {}
    for key, item in raw_fields.items():
        if not isinstance(key, str) or not key.strip() or not isinstance(item, dict):
            raise ProfileError("fields 中每项都必须使用非空字段名作为键")
        field_type = item.get("type", "text")
        if field_type not in {"text", "amount", "number", "date"}:
            raise ProfileError(f"字段 {key} 的 type 必须是 text、amount、number 或 date")
        label = item.get("label", key)
        if not isinstance(label, str) or not label.strip():
            raise ProfileError(f"字段 {key} 的 label 必须是非空字符串")
        aliases = item.get("aliases", [label])
        fields[key] = FieldSpec(
            name=key,
            label=label.strip(),
            aliases=_string_list(aliases, f"fields.{key}.aliases"),
            field_type=field_type,
        )

    if identifier_field in fields and fields[identifier_field].field_type != "text":
        raise ProfileError("identifier.field 若同时出现在 fields 中，其类型必须为 text")

    raw_checks = data.get("checks", [])
    if not isinstance(raw_checks, list):
        raise ProfileError("checks 必须是数组")
    checks: list[dict[str, Any]] = []
    for index, item in enumerate(raw_checks):
        if not isinstance(item, dict):
            raise ProfileError(f"checks[{index}] 必须是对象")
        check_type = item.get("type")
        if check_type not in {"consistent", "required", "date_order"}:
            raise ProfileError(f"checks[{index}].type 不受支持：{check_type}")
        checks.append(dict(item))

    coverage_severity = data.get("coverage_severity", "warning")
    if coverage_severity not in {"error", "warning", "info"}:
        raise ProfileError("coverage_severity 必须是 error、warning 或 info")
    auto_consistent = data.get("auto_consistent_fields", True)
    if not isinstance(auto_consistent, bool):
        raise ProfileError("auto_consistent_fields 必须是布尔值")

    raw_tolerances = data.get("tolerances", {})
    if not isinstance(raw_tolerances, dict):
        raise ProfileError("tolerances 必须是对象")
    tolerances: dict[str, str] = {}
    for key, value in raw_tolerances.items():
        if isinstance(value, bool) or not isinstance(value, (str, int, float)):
            raise ProfileError(f"tolerances.{key} 必须是数字")
        tolerances[key] = str(value)

    return Profile(
        name=name,
        entity_name=entity_name.strip(),
        identifier_field=identifier_field,
        identifier_aliases=identifier_aliases,
        identifier_pattern=identifier_pattern,
        sources=tuple(sources),
        fields=fields,
        checks=tuple(checks),
        coverage_severity=coverage_severity,
        auto_consistent_fields=auto_consistent,
        tolerances=tolerances,
        base_dir=profile_path.parent,
        path=profile_path,
        raw=data,
    )
