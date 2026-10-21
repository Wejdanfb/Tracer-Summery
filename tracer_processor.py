"""
tracer_processor.py
Reads ONE Nursing Tracer workbook (read-only) and extracts:
  - unit name (C6)
  - all checklist items (No., STD No., description, T.M, score, notes)
  - the overall Final Score
  - significant findings (items scored 0 or 1)

Structure observed in the MCH tracer files (and handled dynamically):
  * Sheet "Tracer Checklist" (a helper sheet "Lists" also exists)
  * C6 (merged C6:E6)  -> unit / tracer area
  * H6                 -> formula producing text "Final score: 92.37%"
  * Header row         -> No | STD No. | Checklist items | T.M | Scoring (merged E:G) | Notes
  * Section rows       -> e.g. "NURSING CHAPTER" (no item number) - skipped
  * Item rows          -> numeric No. in column A, score in E (merged E:G), notes in H
  * "Total" row        -> end of checklist
"""
import os
import re
from dataclasses import dataclass, field
from typing import List, Optional

import openpyxl

from excel_utils import (MergedLookup, is_blank, norm_text, normalize_score,
                         parse_item_number, parse_percent, fmt_pct)

UNIT_CELL = "C6"

# Header keywords -> logical column
HEADER_PATTERNS = {
    "no": [r"^no\.?$", r"^n\.?o\.?$", r"^#$", r"^s\.?n\.?$", r"^item no\.?$", r"^serial"],
    "std": [r"^std", r"standard"],
    "item": [r"checklist", r"^items?$", r"description", r"criteria"],
    "tm": [r"^t\.?\s?m\.?$", r"methodology"],
    "score": [r"^scor", r"^score"],
    "notes": [r"^notes?$", r"remarks?", r"comments?", r"^findings?$"],
}

END_MARKERS = ("total", "additional notes", "recommendation", "conducted by", "reviewed by")
FINAL_SCORE_PATTERNS = [r"final\s*score", r"overall\s*score", r"total\s*score", r"tracer\s*score"]


@dataclass
class ChecklistItem:
    row: int
    number: int
    std_no: str
    description: str
    tm: str
    raw_score: object
    score: object  # 0 / 1 / 2 / 'NA' / None / 'INVALID:...'
    notes: Optional[str]

    @property
    def has_notes(self):
        return not is_blank(self.notes)

    @property
    def finding_text(self):
        """Notes exactly as written if present, otherwise the checklist description."""
        if self.has_notes:
            return self.notes if isinstance(self.notes, str) else str(self.notes)
        return self.description

    @property
    def std_label(self):
        """STD No. on one line. Line breaks inside the STD cell come from the
        original Word/PDF layout (e.g. 'QM.14/QM\\n.15', 'HR.3.7\\nHR.4.4').
        A break that splits one code is joined ('QM.14/QM.15'); a break between
        two codes becomes a space ('HR.3.7 HR.4.4'). Nothing else is changed."""
        s = (self.std_no or "").replace("\r", "")
        parts = [p.strip() for p in s.split("\n") if p.strip()]
        out = ""
        for p in parts:
            if not out:
                out = p
            elif out.endswith(("/", ".", "-", ",")) or p.startswith((".", "/", "-", ",")):
                out += p
            else:
                out += " " + p
        return out

    @property
    def summary_text(self):
        """Text written to 'Significant Findings':  STD No. — Notes  (or — description)."""
        std = self.std_label
        return "%s — %s" % (std, self.finding_text) if std else self.finding_text


@dataclass
class TracerResult:
    path: str
    sheet_name: str = ""
    unit: str = ""
    final_score: Optional[float] = None
    score_source: str = ""  # "file" or "computed"
    header_row: int = 0
    columns: dict = field(default_factory=dict)
    items: List[ChecklistItem] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    @property
    def file_name(self):
        return os.path.basename(self.path)

    @property
    def ok(self):
        return not self.errors

    @property
    def findings(self):
        return [it for it in self.items if it.score in (0, 1)]

    def count(self, value):
        return sum(1 for it in self.items if it.score == value)


# ---------------------------------------------------------------------------
def _match(text, key):
    return any(re.search(p, text) for p in HEADER_PATTERNS[key])


def find_header(ws, look=MergedLookup, max_scan_rows=80, max_scan_cols=30):
    """Locate the checklist header row. Returns (row, {logical: col}) or (None, {})."""
    ml = look(ws)
    for r in range(1, min(ws.max_row, max_scan_rows) + 1):
        cols = {}
        for c in range(1, min(ws.max_column, max_scan_cols) + 1):
            # only consider top-left of merged ranges (avoid duplicates)
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            t = norm_text(v)
            if not t:
                continue
            for key in ("no", "std", "item", "tm", "score", "notes"):
                if key not in cols and _match(t, key):
                    # "item" pattern "checklist" must not grab the title row
                    cols[key] = c
                    break
        if {"no", "item", "score"} <= set(cols):
            return r, cols
    return None, {}


def _pick_sheet(wb):
    """Return (ws, header_row, cols) of the sheet that contains the checklist."""
    sheets = list(wb.worksheets)
    sheets.sort(key=lambda s: 0 if re.search(r"tracer|checklist", s.title, re.I) else 1)
    for ws in sheets:
        hr, cols = find_header(ws)
        if hr:
            return ws, hr, cols
    return None, None, {}


def _read_items(ws, header_row, cols, result):
    ml = MergedLookup(ws)
    items = []
    gap = 0
    first_col = min(cols.values())
    for r in range(header_row + 1, ws.max_row + 1):
        num = parse_item_number(ml.value(r, cols["no"]))
        lead = norm_text(ml.value(r, first_col))
        if num is None:
            if items and any(lead.startswith(m) for m in END_MARKERS):
                break
            gap += 1
            if items and gap > 25:
                break
            continue  # blank row or merged section header (e.g. "NURSING CHAPTER")
        gap = 0
        desc = ml.value(r, cols["item"])
        std = ml.value(r, cols["std"]) if "std" in cols else None
        if is_blank(desc) and is_blank(std):
            continue
        raw = ml.value(r, cols["score"])
        notes = ml.value(r, cols["notes"]) if "notes" in cols else None
        tm = ml.value(r, cols["tm"]) if "tm" in cols else None
        items.append(ChecklistItem(
            row=r, number=num,
            std_no="" if std is None else str(std),
            description="" if desc is None else str(desc),
            tm="" if tm is None else str(tm),
            raw_score=raw, score=normalize_score(raw), notes=notes,
        ))
    return items


def _find_final_score(ws_formula, ws_values):
    """Locate the overall Final Score.

    Returns (fraction or None, description_of_location, had_formula_without_cache)
    """
    candidates = []
    no_cache = False
    for row in ws_formula.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str):
                continue
            t = v.lower()
            prio = None
            for i, p in enumerate(FINAL_SCORE_PATTERNS):
                if re.search(p, t):
                    prio = i
                    break
            if prio is None:
                continue
            cached = ws_values.cell(row=cell.row, column=cell.column).value
            is_formula = v.startswith("=")
            # 1) value inside the same cell (e.g. "Final score: 92.37%")
            if is_formula and cached is None:
                no_cache = True
            src_text = cached if is_formula else v
            if isinstance(src_text, str) and re.search(r"\d", src_text.split(":")[-1]):
                pct = parse_percent(src_text.split(":", 1)[-1] if ":" in src_text else src_text)
                if pct is not None:
                    candidates.append((prio, pct, cell.coordinate))
                    continue
            if is_formula:
                continue
            # 2) label cell -> value in the next non-empty cells to the right / below
            neighbours = [(cell.row, cell.column + k) for k in range(1, 8)] + [(cell.row + 1, cell.column)]
            for (r, c) in neighbours:
                fv = ws_formula.cell(row=r, column=c)
                vv = ws_values.cell(row=r, column=c).value
                if fv.value is None:
                    continue
                if isinstance(fv.value, str) and fv.value.startswith("=") and vv is None:
                    no_cache = True
                    break
                pct = parse_percent(vv, fv.number_format)
                if pct is not None:
                    candidates.append((prio, pct, fv.coordinate))
                break
    if candidates:
        candidates.sort(key=lambda x: x[0])
        prio, pct, coord = candidates[0]
        return pct, coord, no_cache
    return None, None, no_cache


def compute_score(items):
    """Same rule as the tracer's own H6 formula:
    (2*count(2) + 1*count(1)) / ((items - NA) * 2)."""
    n2 = sum(1 for i in items if i.score == 2)
    n1 = sum(1 for i in items if i.score == 1)
    na = sum(1 for i in items if i.score == "NA")
    denom = (len(items) - na) * 2
    if denom <= 0:
        return None
    return (n2 * 2 + n1) / denom


# ---------------------------------------------------------------------------
def process_tracer(path):
    res = TracerResult(path=path)
    try:
        # read_only=False so merged ranges are available; files are never saved.
        wb_f = openpyxl.load_workbook(path, data_only=False)
        wb_v = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:  # noqa
        res.errors.append("Could not open workbook (%s)." % e)
        return res

    ws, header_row, cols = _pick_sheet(wb_f)
    if ws is None:
        try:
            from summary_writer import pick_template_sheet
            pick_template_sheet(wb_f)
            res.errors.append("This is a Tracer SUMMARY file, not a unit tracer - skipped.")
        except Exception:
            res.errors.append("Could not find checklist table (header row with No / Checklist items / Scoring).")
        return res
    ws_v = wb_v[ws.title]
    res.sheet_name, res.header_row, res.columns = ws.title, header_row, cols

    if "notes" not in cols:
        res.warnings.append("Notes column could not be detected - checklist descriptions will be used.")
    if "std" not in cols:
        res.warnings.append("STD No. column could not be detected.")

    # --- Unit (C6, merged-aware) ---
    unit = MergedLookup(ws_v).value(6, 3)
    if is_blank(unit):
        res.errors.append("Unit name missing in %s." % UNIT_CELL)
    else:
        res.unit = str(unit).strip()

    # --- Checklist ---
    res.items = _read_items(ws_v, header_row, cols, res)
    if not res.items:
        res.errors.append("Could not find checklist items below the header (row %d)." % header_row)
        return res
    nums = [i.number for i in res.items]
    if len(set(nums)) != len(nums):
        res.warnings.append("Duplicate checklist item numbers found.")
    for it in res.items:
        if it.score in (0, 1) and not it.std_label:
            res.warnings.append("Item %d (row %d) is scored %s but has no STD No." % (it.number, it.row, it.score))
        if it.score is None:
            res.warnings.append("Item %d (row %d) has no score - treated as not scored (excluded)." % (it.number, it.row))
        elif isinstance(it.score, str) and it.score.startswith("INVALID"):
            res.warnings.append("Item %d (row %d) has an unrecognised score %r - excluded." % (it.number, it.row, it.raw_score))

    # --- Final score ---
    pct, where, no_cache = _find_final_score(ws, ws_v)
    computed = compute_score(res.items)
    if pct is not None:
        res.final_score, res.score_source = round(pct, 6), "file (%s)" % where
        if computed is not None and abs(round(computed, 4) - round(pct, 4)) > 0.00011:
            res.warnings.append("Final score in file (%s) differs from the checklist scores (%s). File value used."
                                % (fmt_pct(pct), fmt_pct(computed)))
    elif computed is not None:
        res.final_score, res.score_source = round(computed, 4), "computed"
        if no_cache:
            res.warnings.append("Final Score formula has no saved result (file was not last saved by Excel). "
                                "Score computed with the tracer's own formula: %s." % fmt_pct(computed))
        else:
            res.warnings.append("Could not identify Final Score cell - computed from checklist: %s." % fmt_pct(computed))
    else:
        res.errors.append("Could not identify Final Score.")
    return res
