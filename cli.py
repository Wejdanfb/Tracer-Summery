"""
Tracer Summary - optional COMMAND-LINE mode (the normal way to use the tool is run.bat,
which opens the web interface).

CLI:  python cli.py --template "TRACER SUMMARY TEMPLATE.xlsx" --out "TRACER SUMMARY.xlsx" tracer1.xlsx tracer2.xlsx ...
      optional: --month 09/2026   --keep-duplicates
"""
import argparse
import datetime as dt
import os
import sys

from core import find_duplicate_files, remove_indices, process_all, duplicate_units, write_and_validate
from summary_writer import LAYOUT_PER_FINDING, LAYOUT_PER_UNIT

APP_TITLE = "Tracer Summary Generator — MCH Nursing"
DEFAULT_OUTPUT = "TRACER SUMMARY.xlsx"


def parse_month(text):
    text = (text or "").strip()
    if not text:
        return None
    for fmt in ("%m/%Y", "%m-%Y", "%Y-%m", "%b %Y", "%B %Y", "%m/%y"):
        try:
            d = dt.datetime.strptime(text, fmt)
            return dt.datetime(d.year, d.month, 1)
        except ValueError:
            pass
    raise ValueError("Month/Year '%s' not understood. Use MM/YYYY, e.g. 09/2026." % text)


def same_file(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def final_report(log, ok, failed, output, checks):
    total_findings = sum(len(r.findings) for r in ok)
    log("")
    log("=" * 60)
    all_passed = checks is not None and all(p for p, _ in checks)
    if ok and all_passed and not failed:
        log("Completed successfully.")
    elif ok:
        log("Completed with issues (see above).")
    else:
        log("FAILED — no tracer file could be processed.")
    log("%d files processed." % len(ok))
    log("%d findings extracted." % total_findings)
    if failed:
        log("%d file(s) FAILED:" % len(failed))
        for r in failed:
            log("   ✗ %s — %s" % (r.file_name, "; ".join(r.errors)))
    if checks is not None:
        log("Validation: %d/%d checks passed." % (sum(1 for p, _ in checks if p), len(checks)))
    if output and ok:
        log("Output: %s" % output)
    log("=" * 60)


# ===========================================================================
# CLI
# ===========================================================================
def run_cli(argv):
    try:  # Windows consoles may not support the check-mark characters
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Generate TRACER SUMMARY.xlsx from Nursing Tracer files")
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", default=DEFAULT_OUTPUT)
    ap.add_argument("--month", default="")
    ap.add_argument("--keep-duplicates", action="store_true")
    ap.add_argument("tracers", nargs="+")
    a = ap.parse_args(argv)

    import glob
    paths = []
    for t in a.tracers:  # Windows cmd does not expand wildcards itself
        paths.extend(sorted(glob.glob(t)) if any(ch in t for ch in "*?[") else [t])
    for p in paths:
        if same_file(p, a.out) or same_file(a.template, a.out):
            print("ERROR: output file must not overwrite a source file or the template.")
            return 2
    dups = find_duplicate_files(paths)
    for _, d, orig in dups:
        print("WARNING: Duplicate file detected: %s (same as %s)" % (os.path.basename(d), os.path.basename(orig)))
    if dups and not a.keep_duplicates:
        paths = remove_indices(paths, dups)
        print("         Duplicates removed (use --keep-duplicates to keep them).")
    ok, failed = process_all(paths, log=print)
    for unit, rs in duplicate_units(ok).items():
        print("WARNING: Unit '%s' appears in several files: %s" % (unit, ", ".join(r.file_name for r in rs)))
    checks = None
    if ok:
        _, checks = write_and_validate(a.template, ok, a.out, layout=LAYOUT_PER_UNIT, month=parse_month(a.month), log=print)
    final_report(print, ok, failed, a.out, checks)
    return 0 if ok and not failed and all(p for p, _ in checks) else 1


if __name__ == "__main__":
    sys.exit(run_cli(sys.argv[1:]))
