"""Data processing module — CSV ingestion, type detection, and analysis."""

from __future__ import annotations

import csv
import io
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

# Keywords that hint a column holds currency values
CURRENCY_KEYWORDS: set[str] = {
    "revenue", "cost", "price", "amount", "salary", "wage", "fee",
    "total", "budget", "expense", "income", "profit", "loss",
    "payment", "balance", "tax", "discount", "charge",
}

DATE_FORMATS: list[str] = [
    "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y",
    "%Y/%m/%d", "%d-%b-%Y", "%d %b %Y", "%b %d, %Y",
    "%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S",
    "%d-%m-%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
]


def detect_delimiter(sample: str) -> str:
    """Sniff the CSV delimiter from a text sample."""
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        return dialect.delimiter
    except csv.Error:
        return ","


def detect_encoding(file_path: Path) -> str:
    """Try common encodings and return the first that works."""
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            with open(file_path, encoding=enc) as f:
                f.read(4096)
            return enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    return "utf-8"


def read_csv(source: str | Path) -> tuple[list[str], list[list[str]]]:
    """Read CSV from a file path or raw string.  Returns (headers, rows)."""
    if isinstance(source, Path) or (isinstance(source, str) and Path(source).is_file()):
        path = Path(source)
        encoding = detect_encoding(path)
        logger.info("Reading %s (encoding=%s)", path, encoding)
        with open(path, encoding=encoding, newline="") as f:
            sample = f.read(8192)
            delimiter = detect_delimiter(sample)
            f.seek(0)
            reader = csv.reader(f, delimiter=delimiter)
            rows = list(reader)
    else:
        # Treat source as raw CSV text
        delimiter = detect_delimiter(source[:8192])
        reader = csv.reader(io.StringIO(source), delimiter=delimiter)
        rows = list(reader)

    if not rows:
        raise ValueError("CSV data is empty")

    headers = rows[0]
    data_rows = rows[1:]
    logger.info("Loaded %d rows × %d columns", len(data_rows), len(headers))
    return headers, data_rows


# ── Type detection ───────────────────────────────────────────────────────────

_CURRENCY_RE = re.compile(r"^[\$€£¥₹]?\s*-?[\d,]+\.?\d*$")
_PCT_RE = re.compile(r"^-?[\d,]+\.?\d*\s*%$")
_INT_RE = re.compile(r"^-?[\d,]+$")
_FLOAT_RE = re.compile(r"^-?[\d,]*\.\d+$")


def _parse_date(value: str) -> datetime | None:
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value.strip(), fmt)
        except ValueError:
            continue
    return None


def detect_column_type(header: str, values: list[str]) -> str:
    """Return one of: 'currency', 'percentage', 'date', 'integer', 'float', 'text'."""
    non_empty = [v.strip() for v in values if v.strip()]
    if not non_empty:
        return "text"

    sample = non_empty[:200]  # sample for speed

    header_lower = header.lower().strip()

    # Percentage check
    pct_count = sum(1 for v in sample if _PCT_RE.match(v))
    if pct_count / len(sample) > 0.6:
        return "percentage"

    # Currency check (keyword + pattern)
    if any(kw in header_lower for kw in CURRENCY_KEYWORDS):
        cur_count = sum(1 for v in sample if _CURRENCY_RE.match(v.replace(",", "")))
        if cur_count / len(sample) > 0.5:
            return "currency"

    # Date check
    date_count = sum(1 for v in sample if _parse_date(v) is not None)
    if date_count / len(sample) > 0.6:
        return "date"

    # Numeric checks
    int_count = sum(1 for v in sample if _INT_RE.match(v))
    float_count = sum(1 for v in sample if _FLOAT_RE.match(v))

    if float_count / len(sample) > 0.6:
        # Could also be currency without keyword
        return "float"
    if int_count / len(sample) > 0.6:
        return "integer"

    return "text"


def analyze_columns(
    headers: list[str], rows: list[list[str]]
) -> list[dict[str, Any]]:
    """Return metadata for every column."""
    col_meta: list[dict[str, Any]] = []
    for idx, header in enumerate(headers):
        values = [r[idx] if idx < len(r) else "" for r in rows]
        non_empty = [v for v in values if v.strip()]
        unique = set(non_empty)
        col_type = detect_column_type(header, values)
        col_meta.append({
            "index": idx,
            "header": header,
            "type": col_type,
            "total_count": len(values),
            "null_count": len(values) - len(non_empty),
            "unique_count": len(unique),
            "sample_values": list(unique)[:5],
            "is_categorical": col_type == "text" and 2 <= len(unique) <= 20,
        })
    return col_meta


def coerce_value(raw: str, col_type: str) -> Any:
    """Convert a raw string to a Python value based on detected type."""
    raw = raw.strip()
    if not raw:
        return None
    try:
        if col_type == "currency":
            return float(raw.replace("$", "").replace("€", "").replace("£", "")
                         .replace("¥", "").replace("₹", "").replace(",", "").strip())
        if col_type == "percentage":
            return float(raw.replace("%", "").replace(",", "").strip()) / 100.0
        if col_type == "integer":
            return int(raw.replace(",", ""))
        if col_type == "float":
            return float(raw.replace(",", ""))
        if col_type == "date":
            dt = _parse_date(raw)
            return dt if dt else raw
    except (ValueError, TypeError):
        return raw
    return raw


def find_duplicates(rows: list[list[str]]) -> set[int]:
    """Return 0-based indices of duplicate rows."""
    seen: dict[tuple[str, ...], int] = {}
    dupes: set[int] = set()
    for idx, row in enumerate(rows):
        key = tuple(row)
        if key in seen:
            dupes.add(idx)
            dupes.add(seen[key])
        else:
            seen[key] = idx
    return dupes


def compute_numeric_stats(
    values: list[Any],
) -> dict[str, float | int | None]:
    """Compute sum, avg, min, max for a list of numeric values."""
    nums = [v for v in values if isinstance(v, (int, float))]
    if not nums:
        return {"sum": None, "avg": None, "min": None, "max": None, "count": 0}
    return {
        "sum": sum(nums),
        "avg": sum(nums) / len(nums),
        "min": min(nums),
        "max": max(nums),
        "count": len(nums),
    }
