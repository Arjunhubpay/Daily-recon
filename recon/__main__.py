"""
Command-line entry point.

    python -m recon --root "<share folder>" [--date YYYY-MM-DD] [--lookback 10]

Folder layout expected under --root (one dated sub-folder per day):

    <root>/
        2026-07-21/
            recon-lines-....csv
            transactions_report-....csv
            PaymentHistory-....xlsx
            SettlementReport-....xlsx
            DealHistory-....xlsx
            OpTransactionHistory-....xlsx
            downloaded-statement-....csv
        2026-07-20/
            ...

Outputs (written to --out, default <root>/_recon_reports/<date>/):
    recon_report_<date>.xlsx     full workbook (Summary / Exceptions / Cleared / Matched)
    dashboard_<date>.html        self-contained viewable dashboard
    dashboard_latest.html        copy of the most recent dashboard (for a fixed link)
"""

from __future__ import annotations

import argparse
import datetime as _dt
import sys
from pathlib import Path

from .engine import load_day, reconcile
from .report import write_html, write_xlsx

DATA_EXTS = (".csv", ".xlsx", ".xls")


def _parse_date(s: str) -> _dt.date:
    return _dt.datetime.strptime(s, "%Y-%m-%d").date()


def _day_folder(root: Path, d: _dt.date) -> Path:
    return root / d.isoformat()


def collect_day_map(root: Path, run_date: _dt.date, lookback: int):
    """Load run_date plus each prior day that has a dated sub-folder."""
    day_map = {}
    for offset in range(0, lookback + 1):
        d = run_date - _dt.timedelta(days=offset)
        folder = _day_folder(root, d)
        if not folder.is_dir():
            continue
        files = [p for p in sorted(folder.iterdir())
                 if p.is_file() and p.suffix.lower() in DATA_EXTS]
        if files:
            day_map[d] = load_day(d, files)
    return day_map


def main(argv=None):
    ap = argparse.ArgumentParser(prog="recon", description="Aged daily reconciliation")
    ap.add_argument("--root", required=True,
                    help="Share-folder root containing dated sub-folders (YYYY-MM-DD).")
    ap.add_argument("--date", default=None,
                    help="Run date YYYY-MM-DD (default: today).")
    ap.add_argument("--lookback", type=int, default=10,
                    help="Days to look back when aging breaks (default: 10).")
    ap.add_argument("--out", default=None,
                    help="Output folder (default: <root>/_recon_reports).")
    ap.add_argument("--quiet", action="store_true", help="Suppress the console summary.")
    args = ap.parse_args(argv)

    root = Path(args.root).expanduser()
    if not root.is_dir():
        ap.error(f"--root is not a folder: {root}")

    run_date = _parse_date(args.date) if args.date else _dt.date.today()
    out_root = Path(args.out).expanduser() if args.out else (root / "_recon_reports")

    day_map = collect_day_map(root, run_date, args.lookback)
    if run_date not in day_map:
        print(f"ERROR: no files found for run date {run_date} "
              f"(expected in {_day_folder(root, run_date)})", file=sys.stderr)
        return 2

    result = reconcile(run_date, day_map, args.lookback)

    out_dir = out_root / run_date.isoformat()
    xlsx_path = out_dir / f"recon_report_{run_date.isoformat()}.xlsx"
    html_path = out_dir / f"dashboard_{run_date.isoformat()}.html"
    write_xlsx(result, xlsx_path)
    write_html(result, html_path)
    # a stable "latest" link at the output root
    latest = out_root / "dashboard_latest.html"
    write_html(result, latest)

    if not args.quiet:
        _print_summary(result, xlsx_path, html_path)
    return 0


def _print_summary(result, xlsx_path, html_path):
    print(f"\nRecon Daily — {result['run_date']}  (look-back {result['lookback_days']}d)")
    print(f"  prior days used : {', '.join(result['prior_dates']) or 'none'}")
    print(f"  internal rows   : {result['internal_count']}")
    print(f"  matched         : {len(result['matched'])}")
    print(f"  cleared (aged)  : {len(result['cleared'])}")
    print(f"  exceptions      : {len(result['exceptions'])}")
    print(f"  {'provider':<16}{'match':>7}{'clear':>7}{'excep':>7}")
    for p in result["per_provider"]:
        print(f"  {p['provider']:<16}{p['matched']:>7}{p['cleared']:>7}{p['exceptions']:>7}")
    print(f"\n  report    : {xlsx_path}")
    print(f"  dashboard : {html_path}")
    if result["exceptions"]:
        print(f"\n  {len(result['exceptions'])} exception(s) need manual checking.")


if __name__ == "__main__":
    raise SystemExit(main())
