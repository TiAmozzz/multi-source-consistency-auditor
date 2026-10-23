"""Value normalization used for deterministic cross-source comparison."""

from __future__ import annotations

import re
import unicodedata
from datetime import date
from decimal import Decimal, InvalidOperation

from dateutil import parser as date_parser

_AMOUNT_RE = re.compile(r"[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?")
_DATE_RE = re.compile(
    r"(?P<year>19\d{2}|20\d{2})\s*[-/.年]\s*"
    r"(?P<month>1[0-2]|0?[1-9])\s*[-/.月]\s*"
    r"(?P<day>3[01]|[12]\d|0?[1-9])\s*日?"
)
_COMPACT_DATE_RE = re.compile(r"(?P<year>19\d{2}|20\d{2})(?P<month>1[0-2]|0[1-9])(?P<day>3[01]|[12]\d|0[1-9])")


class NormalizationError(ValueError):
    """Raised when a value cannot be normalized for its declared type."""


def normalize_identifier(raw: str) -> str:
    text = unicodedata.normalize("NFKC", raw or "")
    return re.sub(r"\s+", "", text).upper()


def normalize_text(raw: str) -> str:
    text = unicodedata.normalize("NFKC", raw or "")
    text = re.sub(r"\s+", "", text)
    text = text.strip(" ：:；;，,。.")
    return text.casefold()


def normalize_amount(raw: str) -> str:
    text = unicodedata.normalize("NFKC", raw or "").replace("，", ",")
    match = _AMOUNT_RE.search(text)
    if not match:
        raise NormalizationError("未找到金额数字")

    number = match.group(0).replace(",", "")
    try:
        value = Decimal(number)
    except InvalidOperation as exc:
        raise NormalizationError("金额格式无效") from exc

    if "亿" in text:
        value *= Decimal("100000000")
    elif "万" in text:
        value *= Decimal("10000")

    rendered = format(value.normalize(), "f")
    return rendered.rstrip("0").rstrip(".") if "." in rendered else rendered


def normalize_number(raw: str) -> str:
    text = unicodedata.normalize("NFKC", raw or "").replace(",", "")
    match = _AMOUNT_RE.search(text)
    if not match:
        raise NormalizationError("未找到数字")
    try:
        value = Decimal(match.group(0))
    except InvalidOperation as exc:
        raise NormalizationError("数字格式无效") from exc
    rendered = format(value.normalize(), "f")
    return rendered.rstrip("0").rstrip(".") if "." in rendered else rendered


def normalize_date(raw: str) -> str:
    text = unicodedata.normalize("NFKC", raw or "").strip()
    text = text.replace("年", "-").replace("月", "-").replace("日", "")
    match = _DATE_RE.search(text)
    if not match:
        match = _COMPACT_DATE_RE.search(text)

    try:
        if match:
            parsed = date(
                int(match.group("year")),
                int(match.group("month")),
                int(match.group("day")),
            )
        else:
            parsed = date_parser.parse(text, fuzzy=False).date()
    except (ValueError, OverflowError, TypeError) as exc:
        raise NormalizationError("日期格式无法识别") from exc
    return parsed.isoformat()


def normalize_value(field_type: str, raw: str) -> str:
    if field_type == "text":
        return normalize_text(raw)
    if field_type == "amount":
        return normalize_amount(raw)
    if field_type == "number":
        return normalize_number(raw)
    if field_type == "date":
        return normalize_date(raw)
    raise NormalizationError(f"未知字段类型：{field_type}")


def values_equal(field_type: str, left: str, right: str, tolerance: Decimal = Decimal("0")) -> bool:
    if left == right:
        return True
    if field_type in {"amount", "number"}:
        try:
            return abs(Decimal(left) - Decimal(right)) <= tolerance
        except InvalidOperation:
            return False
    return False


def display_number(value: str) -> str:
    """Return a stable human-readable representation without scientific notation."""
    try:
        return format(Decimal(value), "f")
    except InvalidOperation:
        return value
