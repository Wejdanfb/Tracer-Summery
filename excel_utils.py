"""
excel_utils.py
Low-level helpers for reading/writing Excel with openpyxl:
  - merged-cell aware value reading
  - score normalisation (0 / 1 / 2 / NA)
  - final-score text parsing ("Final score: 92.37%")
  - style copying and row-height estimation
"""
import hashlib
import math
import re
from copy import copy

from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Merged cells
# ---------------------------------------------------------------------------
class MergedLookup:
    """Maps every cell inside a merged range to the range's top-left cell,
    so that a value can be read from ANY visible cell of a merged block."""

    def __init__(self, ws):
        self.ws = ws
        self._map = {}
        for rng in ws.merged_cells.ranges:
            tl = (rng.min_row, rng.min_col)
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    self._map[(r, c)] = tl

    def value(self, row, col):
        r, c = self._map.get((row, col), (row, col))
        return self.ws.cell(row=r, column=c).value

    def is_merged(self, row, col):
        return (row, col) in self._map


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------
def norm_text(v):
    """Lower-case, collapse whitespace, drop trailing ':' / '.' - for header matching."""
    if v is None:
        return ""
    s = str(v).replace(" ", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s.rstrip(":").strip()


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


# ---------------------------------------------------------------------------
# Scores
# ---------------------------------------------------------------------------
NA_TOKENS = {"na", "n/a", "n.a", "n.a.", "not applicable"}


def normalize_score(v):
    """Return one of 0, 1, 2, 'NA', None (blank) or 'INVALID:<raw>'.

    Accepts numbers (1, 1.0), text ("1", " 1 ", "1.0") and NA variants.
    """
    if v is None:
        return None
    if isinstance(v, bool):
        return "INVALID:%r" % v
    if isinstance(v, (int, float)):
        if float(v).is_integer() and int(v) in (0, 1, 2):
            return int(v)
        return "INVALID:%r" % v
    s = str(v).replace(" ", " ").strip()
    if s == "":
        return None
    if s.lower() in NA_TOKENS:
        return "NA"
    try:
        f = float(s)
        if f.is_integer() and int(f) in (0, 1, 2):
            return int(f)
    except ValueError:
        pass
    return "INVALID:%r" % v


def parse_item_number(v):
    """Checklist 'No.' values are integers (1, 2, '3', '4.'). Returns int or None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() and v > 0 else None
    s = str(v).strip().rstrip(".").strip()
    if s.isdigit():
        return int(s)
    return None


_PCT_RE = re.compile(r"(-?\d+(?:[.,]\d+)?)\s*%")
_NUM_RE = re.compile(r"(-?\d+(?:[.,]\d+)?)")


def parse_percent(v, cell_number_format=None):
    """Convert a percentage-like value into a fraction (0.9237).

    - numeric 0.9237 (percent-formatted or <= 1)  -> 0.9237
    - numeric 92.37                                 -> 0.9237
    - text "Final score: 92.37%" / "92.37 %"        -> 0.9237
    Returns None when nothing usable is found.
    """
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        if cell_number_format and "%" in cell_number_format:
            return f
        if 0 <= f <= 1:
            return f
        if 1 < f <= 100:
            return f / 100.0
        return None
    s = str(v)
    m = _PCT_RE.search(s)
    if m:
        return float(m.group(1).replace(",", ".")) / 100.0
    m = _NUM_RE.search(s)
    if m:
        f = float(m.group(1).replace(",", "."))
        if 0 <= f <= 1:
            return f
        if 1 < f <= 100:
            return f / 100.0
    return None


def fmt_pct(fraction):
    return "%.2f%%" % (fraction * 100.0) if fraction is not None else "?"


# ---------------------------------------------------------------------------
# Files
# ---------------------------------------------------------------------------
def file_sha1(path):
    h = hashlib.sha1()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


# ---------------------------------------------------------------------------
# Styles / layout
# ---------------------------------------------------------------------------
def copy_cell_style(src, dst):
    """Copy every style attribute of src onto dst (same workbook)."""
    if src.has_style:
        dst._style = copy(src._style)


def estimate_row_height(texts_and_widths, min_height=30.0, line_height=14.5, padding=8.0,
                        max_height=409.0):
    """Estimate a row height for wrapped text.

    texts_and_widths: iterable of (text, column_width_in_excel_units).
    Excel cannot auto-fit rows written by openpyxl, so we approximate.
    """
    lines_needed = 1
    for text, width in texts_and_widths:
        if text is None:
            continue
        cpl = max(8, int((width or 10) * 1.05))  # characters per line (bold Calibri 11)
        lines = 0
        for para in str(text).split("\n"):
            lines += max(1, math.ceil(len(para.rstrip()) / cpl))
        lines_needed = max(lines_needed, lines)
    h = lines_needed * line_height + padding
    return float(min(max_height, max(min_height, h)))


def col_letter(idx):
    return get_column_letter(idx)
