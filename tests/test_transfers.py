"""Internal-transfer detection: paired legs -> Matched; a debit leg with no
credit leg -> Credit pending. Run:  python tests/test_transfers.py
"""

import datetime as _dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from recon.engine import DayData, internal_transfers  # noqa: E402


def _row(prov, ref, rfao, amt, tt, sd):
    return {"Payment Provider": prov, "Reference": ref, "Ref for Account Owner": rfao,
            "Value Date": "2026-07-21", "Amount": amt, "Currency": "AED",
            "Transaction Type": tt, "Supplementary Details": sd}


def run():
    day = DayData(date=_dt.date(2026, 7, 21))
    for p in ("NBF", "Currency Cloud", "Corpay", "Zand"):
        day.internal[p] = []
    U1 = "f4ae3e2c-2011-42ef-bff0-0e7fbf000001"
    U2 = "2f9f1bb3-d431-4077-b4f6-b2d2c6000002"
    day.internal["Zand"] = [
        # matched pair: both legs share the RFAO UUID, equal amount
        _row("ZAND", "D111_aaaaaaaa", U1, "1858.80", "Debit", "internal transfer DNT"),
        _row("ZAND", "D111_bbbbbbbb", U1, "1858.80", "Credit", "internal transfer DNT"),
        # credit pending: only the debit leg present
        _row("ZAND", "D222_cccccccc", U2, "7000.00", "Debit", "internal transfer DNT"),
        # a normal payout (not internal) — must be ignored
        _row("ZAND", "D333_dddddddd", "somebeneficiary", "500.00", "Debit", "Payout to vendor"),
    ]

    transfers, s = internal_transfers(day)
    assert s["total"] == 2, s
    assert s["matched"] == 1, s
    assert s["credit_pending"] == 1, s

    by_status = {t["Status"]: t for t in transfers}
    assert by_status["Matched"]["Amount"] == "1858.80"
    assert by_status["Matched"]["Debit Ref"] == "D111_aaaaaaaa"
    assert by_status["Matched"]["Credit Ref"] == "D111_bbbbbbbb"
    assert by_status["Credit pending"]["Debit Ref"] == "D222_cccccccc"
    assert by_status["Credit pending"]["Credit Ref"] == ""
    print("test_transfers: OK")


def test_transfers():
    run()


if __name__ == "__main__":
    run()
