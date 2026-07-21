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

from . import ledger
from .engine import internal_transfers, load_day, reconcile
from .report import write_daily_snapshot, write_html

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
    # Default output to an existing "Output" folder if present, else _recon_reports.
    if args.out:
        out_root = Path(args.out).expanduser()
    elif (root / "Output").is_dir():
        out_root = root / "Output"
    else:
        out_root = root / "_recon_reports"

    day_map = collect_day_map(root, run_date, args.lookback)
    if run_date not in day_map:
        print(f"ERROR: no files found for run date {run_date} "
              f"(expected in {_day_folder(root, run_date)})", file=sys.stderr)
        return 2

    result = reconcile(run_date, day_map, args.lookback)
    transfers, tsum = internal_transfers(day_map[run_date])

    # --- rolling open-exceptions ledger (carried forward across runs) --------
    master = out_root / "Open_Exceptions.xlsx"
    open_map, cleared = ledger.load(master)
    lsum = ledger.update(run_date, day_map, args.lookback, result, open_map, cleared)
    ledger.save(master, open_map, cleared)

    # --- dated record in the date folder + viewable dashboard ----------------
    date_folder = _day_folder(root, run_date)
    snapshot = date_folder / f"Recon_Summary_{run_date.isoformat()}.xlsx"
    dated_html = date_folder / f"dashboard_{run_date.isoformat()}.html"
    write_daily_snapshot(result, lsum, snapshot, transfers)
    write_html(result, dated_html, lsum, transfers)
    # stable "latest" dashboard at the output root
    write_html(result, out_root / "dashboard_latest.html", lsum, transfers)

    if not args.quiet:
        _print_summary(result, lsum, master, snapshot, tsum)
    return 0


def _print_summary(result, lsum, master, snapshot, tsum):
    print(f"\nRecon Daily — {result['run_date']}  (look-back {result['lookback_days']}d)")
    print(f"  prior days used   : {', '.join(result['prior_dates']) or 'none'}")
    print(f"  internal rows     : {result['internal_count']}")
    print(f"  matched (same day): {len(result['matched'])}")
    print(f"  cleared vs prior  : {len(result['cleared'])}")
    print(f"  new exceptions    : {lsum['new_today']}")
    print(f"  cleared today     : {lsum['cleared_today']}")
    print(f"  OPEN (total)      : {lsum['open_total']}  (manual review: {lsum['open_manual']})")
    print(f"  internal transfers: {tsum['total']}  "
          f"(matched {tsum['matched']}, credit pending {tsum['credit_pending']}, "
          f"debit pending {tsum['debit_pending']}, mismatch {tsum['amount_mismatch']}, "
          f"unpaired {tsum['unpaired']})")
    print(f"\n  open ledger : {master}")
    print(f"  dated file  : {snapshot}")


if __name__ == "__main__":
    raise SystemExit(main())
