"""
core.py
Orchestration shared by the GUI (main.py) and the command-line mode.
"""
import os
from collections import OrderedDict, defaultdict

from excel_utils import file_sha1, fmt_pct
from tracer_processor import process_tracer
from summary_writer import write_summary, LAYOUT_PER_UNIT
from validator import validate_output


def find_duplicate_files(paths):
    """Return list of (index, dup_path, original_path) for repeated paths or byte-identical files.
    The FIRST occurrence is treated as the original; later ones are the duplicates."""
    dups = []
    seen_path, seen_hash = {}, {}
    for i, p in enumerate(paths):
        key = os.path.normcase(os.path.abspath(p))
        if key in seen_path:
            dups.append((i, p, seen_path[key]))
            continue
        seen_path[key] = p
        try:
            h = file_sha1(p)
        except OSError:
            continue
        if h in seen_hash:
            dups.append((i, p, seen_hash[h]))
        else:
            seen_hash[h] = p
    return dups


def remove_indices(paths, dups):
    drop = {i for i, _, _ in dups}
    return [p for i, p in enumerate(paths) if i not in drop]


def process_all(paths, log=print):
    ok, failed = [], []
    for p in paths:
        name = os.path.basename(p)
        log("Processing %s..." % name)
        res = process_tracer(p)
        for w in res.warnings:
            log("  WARNING: %s — %s" % (name, w))
        if not res.ok:
            for e in res.errors:
                log("  ERROR: %s — %s" % (name, e))
            failed.append(res)
            continue
        log("  ✓ Sheet: %s (checklist header row %d, %d items)" % (res.sheet_name, res.header_row, len(res.items)))
        log("  ✓ Unit detected: %s" % res.unit)
        log("  ✓ Final score detected: %s  [%s]" % (fmt_pct(res.final_score), res.score_source))
        log("  ✓ Scores: 2=%d, 1=%d, 0=%d, NA=%d" % (res.count(2), res.count(1), res.count(0), res.count("NA")))
        log("  ✓ %d findings extracted" % len(res.findings))
        ok.append(res)
    return ok, failed


def duplicate_units(results):
    by_unit = defaultdict(list)
    for r in results:
        by_unit[r.unit.strip().upper()].append(r)
    return OrderedDict((k, v) for k, v in by_unit.items() if len(v) > 1)


def write_and_validate(template, results, output, layout=LAYOUT_PER_UNIT, month=None, log=print):
    folder = os.path.dirname(os.path.abspath(output))
    os.makedirs(folder, exist_ok=True)
    info = write_summary(template, results, output, layout=layout, month_value=month, log=log)
    log("Saved: %s" % output)
    log("")
    log("Validating output...")
    checks = validate_output(output, template, results, info)
    for passed, msg in checks:
        log(("  ✓ " if passed else "  ✗ FAILED: ") + msg)
    return info, checks
