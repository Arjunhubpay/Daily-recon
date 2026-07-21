"""
Rolling open-exceptions ledger.

A single master workbook (`Open_Exceptions.xlsx`) is carried forward across
runs. Each day:

  * every currently-open item is re-checked against the recent window — if the
    counterparty has since reported it, the item is CLEARED and moved to the
    cleared log (stamped with its original First Seen date);
  * today's new hard exceptions (unreconciled even after the look-back) are
    ADDED as open, dated with today's First Seen;
  * an item still open beyond the look-back window is flagged "Manual Review".

So the "Open Exceptions" sheet is always the live list of what is currently
outstanding, and "Cleared Log" is the audit trail of what has resolved.
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path

from .engine import _fmt_amt, amt_key, reconciles_in_window

OPEN_COLS = [
    "Exception Key", "First Seen", "Provider", "Side", "Internal Ref",
    "Provider Ref", "Value Date", "Amount", "Currency", "Days Open",
    "Status", "Last Checked", "Description",
]
CLEARED_COLS = [
    "Exception Key", "First Seen", "Cleared Date", "Days To Clear", "Provider",
    "Side", "Internal Ref", "Provider Ref", "Amount", "Currency", "Cleared Against",
]

SHEET_OPEN = "Open Exceptions"
SHEET_CLEARED = "Cleared Log"


def exc_key(rec):
    """Stable identity for an exception across runs."""
    amt = rec.amount_internal if rec.amount_internal is not None else rec.amount_provider
    return f"{rec.provider}|{rec.side}|{(rec.provider_ref or rec.internal_ref or '').strip()}|{_fmt_amt(amt_key(amt))}"


def _parse_date(s):
    if not s:
        return None
    try:
        return _dt.datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Load / save
# ---------------------------------------------------------------------------
def load(path: Path):
    """Return (open_map: {key: row_dict}, cleared_rows: [row_dict])."""
    open_map, cleared = {}, []
    if not path.exists():
        return open_map, cleared
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if SHEET_OPEN in wb.sheetnames:
        ws = wb[SHEET_OPEN]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            hdr = [str(h) if h is not None else "" for h in rows[0]]
            for r in rows[1:]:
                d = {hdr[i]: ("" if v is None else v) for i, v in enumerate(r)}
                if d.get("Exception Key"):
                    open_map[str(d["Exception Key"])] = d
    if SHEET_CLEARED in wb.sheetnames:
        ws = wb[SHEET_CLEARED]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            hdr = [str(h) if h is not None else "" for h in rows[0]]
            for r in rows[1:]:
                d = {hdr[i]: ("" if v is None else v) for i, v in enumerate(r)}
                if d.get("Exception Key"):
                    cleared.append(d)
    wb.close()
    return open_map, cleared


def save(path: Path, open_map: dict, cleared: list):
    import openpyxl
    from openpyxl.styles import Font, PatternFill

    wb = openpyxl.Workbook()
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="B5533A")

    def write(ws, cols, rows):
        ws.append(cols)
        for c in ws[1]:
            c.font = head_font
            c.fill = head_fill
        for row in rows:
            ws.append([row.get(c, "") for c in cols])
        ws.freeze_panes = "A2"
        for i, col in enumerate(cols, 1):
            lens = [len(col)] + [len(str(row.get(col, ""))) for row in rows]
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = min(max(lens) + 3, 60)

    ws_open = wb.active
    ws_open.title = SHEET_OPEN
    # newest first, manual-review items on top
    ordered = sorted(open_map.values(),
                     key=lambda d: (d.get("Status") != "Manual Review", str(d.get("First Seen", ""))))
    write(ws_open, OPEN_COLS, ordered)
    write(wb.create_sheet(SHEET_CLEARED), CLEARED_COLS, cleared)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ---------------------------------------------------------------------------
# Daily update
# ---------------------------------------------------------------------------
def update(run_date: _dt.date, day_map: dict, lookback_days: int, result: dict,
           open_map: dict, cleared: list):
    """Mutates open_map/cleared in place. Returns a summary dict for the day."""
    window = [run_date - _dt.timedelta(days=n) for n in range(0, lookback_days + 1)]

    cleared_today = []

    # 1) re-check every currently-open item against the recent window
    for key, row in list(open_map.items()):
        hit = reconciles_in_window(
            row.get("Provider"), row.get("Side"), row.get("Provider Ref"),
            row.get("Amount"), row.get("Currency"), day_map, window)
        if hit is not None:
            first = _parse_date(row.get("First Seen")) or run_date
            crow = {
                "Exception Key": key,
                "First Seen": row.get("First Seen"),
                "Cleared Date": run_date.isoformat(),
                "Days To Clear": (run_date - first).days,
                "Provider": row.get("Provider"),
                "Side": row.get("Side"),
                "Internal Ref": row.get("Internal Ref"),
                "Provider Ref": row.get("Provider Ref"),
                "Amount": row.get("Amount"),
                "Currency": row.get("Currency"),
                "Cleared Against": hit.isoformat(),
            }
            cleared.append(crow)
            cleared_today.append(crow)
            del open_map[key]
        else:
            first = _parse_date(row.get("First Seen")) or run_date
            days_open = (run_date - first).days
            row["Days Open"] = days_open
            row["Last Checked"] = run_date.isoformat()
            row["Status"] = "Manual Review" if days_open > lookback_days else "Open"

    # 2) add today's new hard exceptions
    new_today = []
    for rec in result["exceptions"]:
        key = exc_key(rec)
        if key in open_map:
            open_map[key]["Last Checked"] = run_date.isoformat()
            continue
        amt = rec.amount_internal if rec.amount_internal is not None else rec.amount_provider
        row = {
            "Exception Key": key,
            "First Seen": run_date.isoformat(),
            "Provider": rec.provider,
            "Side": rec.side,
            "Internal Ref": rec.internal_ref,
            "Provider Ref": rec.provider_ref,
            "Value Date": rec.value_date,
            "Amount": _fmt_amt(amt_key(amt)),
            "Currency": rec.currency,
            "Days Open": 0,
            "Status": "Open",
            "Last Checked": run_date.isoformat(),
            "Description": rec.description,
        }
        open_map[key] = row
        new_today.append(row)

    return {
        "open_total": len(open_map),
        "open_manual": sum(1 for r in open_map.values() if r.get("Status") == "Manual Review"),
        "new_today": len(new_today),
        "cleared_today": len(cleared_today),
        "cleared_today_rows": cleared_today,
        "new_today_rows": new_today,
        "open_rows": list(open_map.values()),
    }
