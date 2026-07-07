from __future__ import annotations
import datetime as dt
import re

_QUOTED = re.compile(r'"([^"]*)"')
_PLACEHOLDER = re.compile(r"[0#]")
_URL = re.compile(r"^https?://([^/\s]+)(/\S*)?$", re.IGNORECASE)
# Excel formula-error literals: blanked so broken source formulas don't show to a
# reader. They mark a formula problem in the sheet, not real data.
_EXCEL_ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!",
                 "#NULL!", "#SPILL!", "#CALC!", "#GETTING_DATA"}


def _shorten_url(s: str) -> str:
    """Collapse a long bare URL to 'domain/…/tail' so a cell full of pasted Google
    Docs links reads cleanly instead of dumping 80 characters of raw URL."""
    t = s.strip()
    if len(t) <= 50:
        return s
    m = _URL.match(t)
    if not m:
        return s
    domain, path = m.group(1), (m.group(2) or "")
    tail = path.rstrip("/").split("/")[-1].split("?")[0][:18]
    return f"{domain}/…/{tail}" if tail else domain


def format_cell_value(value, number_format: str | None) -> str:
    """Render a cell's display string from its value and Excel number format.

    Handles the formats that show up in real process sheets: decimals, percent,
    thousands, scientific (0.00E+00), quoted-literal units ('0" °C"'), currency
    and accounting negatives, and date/time (including time-of-day). Anything
    unrecognised falls back to the plain value so nothing is ever lost."""
    if value is None:
        return ""
    if isinstance(value, str):
        if value.strip() in _EXCEL_ERRORS:
            return ""
        return _shorten_url(value)
    if isinstance(value, dt.datetime):
        return _format_datetime(value, number_format or "")
    if isinstance(value, dt.time):
        return value.strftime("%H:%M:%S" if _has_seconds(number_format or "") else "%H:%M")
    if isinstance(value, dt.date):
        return value.strftime("%Y-%m-%d")
    if not isinstance(value, (int, float)):
        return str(value)

    fmt = (number_format or "General").strip()
    if fmt in ("General", "@", ""):
        return _plain_number(value)

    section = _pick_section(fmt, value)
    core = _QUOTED.sub("", section)          # format grammar without literal text
    if not _PLACEHOLDER.search(core):        # e.g. "weird" — not a real number format
        return _plain_number(value)

    use_parens = value < 0 and "(" in section
    v = abs(value) if use_parens else value
    prefix, suffix = _affixes(section)

    try:
        if re.search(r"[eE][+-]?0", core):
            body = f"{v:.{_decimals(core)}E}"
        elif "%" in core:
            body = f"{v * 100:.{_decimals(core)}f}%"
        else:
            dec = _decimals(core)
            body = f"{v:,.{dec}f}" if "," in core else f"{v:.{dec}f}"
    except (ValueError, TypeError):
        return _plain_number(value)

    text = f"{prefix}{body}{suffix}"
    return f"({text})" if use_parens else text


def _plain_number(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _pick_section(fmt: str, value) -> str:
    """Excel formats are 'positive;negative;zero;text' sections. Pick the one that
    applies, stripping [Red]/[$..]/locale color-and-condition brackets."""
    sections = _split_sections(fmt)
    sec = sections[0]
    if value < 0 and len(sections) > 1 and sections[1].strip():
        sec = sections[1]
    elif value == 0 and len(sections) > 2 and sections[2].strip():
        sec = sections[2]
    return re.sub(r"\[[^\]]*\]", "", sec).strip()


def _split_sections(fmt: str) -> list[str]:
    out, cur, in_q = [], "", False
    for ch in fmt:
        if ch == '"':
            in_q = not in_q
            cur += ch
        elif ch == ";" and not in_q:
            out.append(cur); cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


def _affixes(section: str) -> tuple[str, str]:
    """Literal text (quoted segments, $, unit suffixes) before/after the number."""
    m = _PLACEHOLDER.search(section)
    if not m:
        return "", ""
    last = max(p.start() for p in _PLACEHOLDER.finditer(section))
    left, right = section[:m.start()], section[last + 1:]

    def lit(s: str) -> str:
        quoted = "".join(_QUOTED.findall(s))
        rest = _QUOTED.sub("", s)
        rest = rest.replace("\\", "").replace("_", " ").replace("*", "")
        rest = re.sub(r"[0#,.%eE?+\-()]", "", rest)   # drop grammar, keep $ and units
        return (quoted + rest) if quoted and not rest else (rest + quoted)

    return lit(left).strip(), (" " + lit(right).strip()).rstrip() if lit(right).strip() else ""


def _decimals(fmt: str) -> int:
    m = re.search(r"\.(0+)", fmt)
    return len(m.group(1)) if m else 0


def _has_seconds(fmt: str) -> bool:
    return "s" in fmt.lower()


def _format_datetime(value: dt.datetime, fmt: str) -> str:
    f = fmt.lower()
    has_date = ("y" in f) or ("d" in f)
    has_time = ("h" in f) or ("s" in f)
    if has_time:
        spec = "%H:%M:%S" if "s" in f else "%H:%M"
        return value.strftime(("%Y-%m-%d " + spec) if has_date else spec)
    return value.strftime("%Y-%m-%d")
