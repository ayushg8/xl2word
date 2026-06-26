from __future__ import annotations
import datetime as dt
import re


def format_cell_value(value, number_format: str | None) -> str:
    if value is None:
        return ""
    fmt = (number_format or "General").strip()
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%Y-%m-%d")
    if fmt in ("General", "@", ""):
        return str(value)
    if isinstance(value, (int, float)):
        try:
            if "%" in fmt:
                decimals = _decimals(fmt)
                return f"{value * 100:.{decimals}f}%"
            if "," in fmt and "." not in fmt:
                return f"{value:,.0f}"
            if "." in fmt:
                return f"{value:.{_decimals(fmt)}f}"
            if re.fullmatch(r"0+", fmt):
                return f"{int(round(value))}"
        except (ValueError, TypeError):
            return str(value)
    return str(value)


def _decimals(fmt: str) -> int:
    m = re.search(r"\.(0+)", fmt)
    return len(m.group(1)) if m else 0
