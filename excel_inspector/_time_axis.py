"""Small time-axis recognizers for wide sparse table heuristics (issue #22)."""

from __future__ import annotations

import datetime as _dt
import re

_DATE_LIKE_RE = re.compile(
    r"^\d{4}(?:[-/.]\d{1,2}(?:[-/.]\d{1,2})?)?$"
)
_QUARTER_LIKE_RE = re.compile(
    r"^(?:\d{4}\s*[-/]?\s*[Qq][1-4]|[Qq][1-4]\s*[-/]?\s*\d{4})$"
)
_KOREAN_QUARTER_RE = re.compile(r"^\d{4}\s*년\s*[1-4]\s*분기$")

_TIME_AXIS_LABELS = {
    "date",
    "day",
    "month",
    "period",
    "quarter",
    "time",
    "year",
    "기간",
    "날짜",
    "년",
    "년도",
    "년월",
    "분기",
    "시점",
    "연도",
    "월",
    "일자",
}


def _clean_text(value: str) -> str:
    """Normalize a string enough for conservative time-axis matching."""

    return value.strip().strip(":").strip().lower()


def is_time_axis_label(value: object) -> bool:
    """Whether ``value`` names a date/period axis column."""

    if not isinstance(value, str):
        return False
    text = _clean_text(value)
    return text in _TIME_AXIS_LABELS


def is_time_axis_value(value: object) -> bool:
    """Whether ``value`` looks like a date/period coordinate, not a series code."""

    if value is None or (isinstance(value, str) and value == ""):
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return True
    if isinstance(value, int):
        return 1000 <= value <= 9999
    if isinstance(value, float):
        return value.is_integer() and 1000 <= value <= 9999
    if not isinstance(value, str):
        return False

    text = _clean_text(value)
    return (
        _DATE_LIKE_RE.match(text) is not None
        or _QUARTER_LIKE_RE.match(text) is not None
        or _KOREAN_QUARTER_RE.match(text) is not None
    )
