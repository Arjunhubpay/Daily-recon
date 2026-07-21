"""Regression test for the recon engine: builds fixtures in a temp folder
covering same-day match, aged clearing, and true exceptions, then asserts the
totals. Run with:  python -m pytest tests/  (or)  python tests/test_engine.py
"""

import csv
import datetime as _dt
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from recon.__main__ import collect_day_map  # noqa: E402
from recon.engine import reconcile  # noqa: E402

INT_H = ["Payment Provider", "Transaction Type", "Reference", "Value Date",
         "Amount", "Currency", "Supplementary Details", "Ref for Account Owner"]
NBF_PRE = [["Bank"], ["Statement"], [""], ["Acct"], ["Period"], [""], [""]]
NBF_H = ["Date", "Value Date", "Description", "Debit", "Credit", "Balance"]
CC_H = ["Reference", "Amount", "Currency", "Status"]
ZAND_H = ["Outgoing Id", "Ref for Account Owner", "Credit", "Debit", "Currency"]


def _w(root, day, name, header, rows, preamble=None):
    d = root / day
    d.mkdir(parents=True, exist_ok=True)
    with open(d / name, "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        for p in (preamble or []):
            wr.writerow(p)
        wr.writerow(header)
        wr.writerows(rows)


def _build(root):
    _w(root, "2026-07-21", "recon-lines.csv", INT_H, [
        ["NBF", "Debit_REMITTANCE", "FT12345ABCDE", "2026-07-21", "1000.00", "AED", "", ""],
        ["NBF", "Debit_REMITTANCE", "FT99999ZZZZZ", "2026-07-21", "2000.00", "AED", "", ""],
        ["NBF", "Debit_REMITTANCE", "FT00000AAAAA", "2026-07-21", "3000.00", "AED", "", ""],
        ["NBF", "Debit_FEE_VAT", "FTFEEFEEFEE0", "2026-07-21", "5.00", "AED", "", ""],
        ["CURRENCY_CLOUD", "Debit_REMITTANCE", "CC-100", "2026-07-21", "150.25", "USD", "", ""],
        ["CURRENCY_CLOUD", "Debit_REMITTANCE", "CC-200", "2026-07-21", "275.50", "EUR", "", ""],
        ["CURRENCY_CLOUD", "Debit_REMITTANCE", "CC-300", "2026-07-21", "400.00", "USD", "", ""],
        ["CURRENCY_CLOUD", "Debit_REMITTANCE", "CC-400", "2026-07-21", "500.00", "GBP", "", ""],
        ["CORPAY", "Debit_REMITTANCE", "", "2026-07-21", "1111.00", "AED", "1001", ""],
        ["CORPAY", "Debit_REMITTANCE", "", "2026-07-21", "1222.00", "AED", "1002", ""],
        ["CORPAY", "Debit_REMITTANCE", "1003", "2026-07-21", "1333.00", "AED", "", ""],
        ["CORPAY", "Debit_REMITTANCE", "", "2026-07-21", "1444.00", "AED", "1004", ""],
        ["CORPAY", "Debit_REMITTANCE", "", "2026-07-21", "1555.00", "AED", "1005", ""],
        ["ZAND", "Debit_REMITTANCE", "ZAND777_deadbeef", "2026-07-21", "777.00", "AED", "", ""],
        ["ZAND", "Debit_REMITTANCE", "NOPE_00000000", "2026-07-21", "888.00", "AED", "", "RFAO888"],
        ["ZAND", "Debit_REMITTANCE", "ZAND999_cafebabe", "2026-07-21", "999.00", "AED", "", ""],
        ["ZAND", "Debit_REMITTANCE", "ZANDXXX_11111111", "2026-07-21", "1212.00", "AED", "", ""],
    ])
    _w(root, "2026-07-21", "OpTransactionHistory.csv", NBF_H, [
        ["2026-07-21", "2026-07-21", "Outward FT12345ABCDE ben X", "1000.00", "", "5000"],
        ["2026-07-21", "2026-07-21", "Computer generated FT12345ABCDE", "", "", "5000"],
        ["2026-07-21", "2026-07-21", "Outward FT55555BBBBB ben Y", "4321.00", "", "4000"],
    ], preamble=NBF_PRE)
    _w(root, "2026-07-21", "transactions_report.csv", CC_H, [
        ["CC-100", "150.25", "USD", "completed"],
        ["CC-XYZ", "275.50", "EUR", "completed"],
        ["CC-EXTRA", "9999.00", "JPY", "completed"],
    ])
    _w(root, "2026-07-21", "PaymentHistory.csv", ["Deal #", "Amount", "Status"],
       [["1001", "1111.00", "Completed"]])
    _w(root, "2026-07-21", "SettlementReport.csv", ["Deal #", "Settlement Amount", "Status"],
       [["1002", "1222.00", "Paid"], ["1099", "8888.00", "Pending"]])
    _w(root, "2026-07-21", "DealHistory.csv", ["Deal #", "Purchased", "Sold"],
       [["1003", "1333.00", "AED"]])
    _w(root, "2026-07-21", "downloaded-statement.csv", ZAND_H, [
        ["ZAND777", "", "", "777.00", "AED"],
        ["RFAO888", "", "", "888.00", "AED"],
        ["ZANDZZZ", "", "", "5000.00", "AED"],
    ])
    # prior days that clear breaks
    _w(root, "2026-07-20", "transactions_report.csv", CC_H, [["CC-300", "400.00", "USD", "x"]])
    _w(root, "2026-07-20", "PaymentHistory.csv", ["Deal #", "Amount", "Status"], [["1004", "1444.00", "C"]])
    _w(root, "2026-07-19", "OpTransactionHistory.csv", NBF_H,
       [["2026-07-19", "2026-07-19", "Outward FT99999ZZZZZ ben Z", "2000.00", "", "3000"]], preamble=NBF_PRE)
    _w(root, "2026-07-18", "downloaded-statement.csv", ZAND_H, [["ZAND999", "", "", "999.00", "AED"]])


def run():
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        _build(root)
        run_date = _dt.date(2026, 7, 21)
        day_map = collect_day_map(root, run_date, 10)

        # internal-driven (default): only internal-not-in-provider exceptions
        ri = reconcile(run_date, day_map, 10)
        assert len(ri["exceptions"]) == 4, len(ri["exceptions"])

        # two-sided: also provider-not-in-internal
        r = reconcile(run_date, day_map, 10, two_sided=True)
        assert r["internal_count"] == 16, r["internal_count"]
        assert len(r["matched"]) == 8, len(r["matched"])
        assert len(r["cleared"]) == 4, len(r["cleared"])
        assert len(r["exceptions"]) == 7, len(r["exceptions"])

        by = {p["provider"]: p for p in r["per_provider"]}
        assert (by["NBF"]["matched"], by["NBF"]["cleared"], by["NBF"]["exceptions"]) == (1, 1, 2)
        assert (by["Currency Cloud"]["matched"], by["Currency Cloud"]["cleared"],
                by["Currency Cloud"]["exceptions"]) == (2, 1, 2)
        assert (by["Corpay"]["matched"], by["Corpay"]["cleared"], by["Corpay"]["exceptions"]) == (3, 1, 1)
        assert (by["Zand"]["matched"], by["Zand"]["cleared"], by["Zand"]["exceptions"]) == (2, 1, 2)

        # cleared items are stamped with the prior date they reconciled against
        cleared_dates = {c.provider_ref: c.cleared_date for c in r["cleared"]}
        assert cleared_dates["FT99999ZZZZZ"] == "2026-07-19"
        assert cleared_dates["CC-300"] == "2026-07-20"
        assert cleared_dates["ZAND999"] == "2026-07-18"
    print("test_engine: OK")


def test_engine():
    run()


if __name__ == "__main__":
    run()
