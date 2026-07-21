"""Internal-transfer detection from statements + credit verification:
a Corpay intercompany payment whose credit posted in NBF -> Matched; one whose
credit has NOT posted -> Credit pending. Run: python tests/test_transfers.py
"""

import datetime as _dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from recon.engine import DayData, internal_transfers  # noqa: E402


def run():
    day = DayData(date=_dt.date(2026, 7, 21))
    # Corpay outgoing intercompany payments to the NBF own account
    day.corpay_pay_rows = [
        {"Deal #": "500", "Amount": "500.00", "Currency": "AED",
         "Purpose Of Payment": "INTERCOMPANY PAYMENT",
         "Identifier": "Hubpay Limited NBF Client Money AED", "Beneficiary": "Hubpay Limited",
         "Reference": "Own account transfer"},                     # credit posts in NBF -> Matched
        {"Deal #": "1400000", "Amount": "1400000.00", "Currency": "AED",
         "Purpose Of Payment": "INTERCOMPANY PAYMENT",
         "Identifier": "Hubpay Limited NBF Client Money AED", "Beneficiary": "Hubpay Limited",
         "Reference": "Own account transfer"},                     # NOT in NBF -> Credit pending
        {"Deal #": "777", "Amount": "300.00", "Currency": "AED",
         "Purpose Of Payment": "BILL PAYMENT", "Identifier": "Some Vendor",
         "Beneficiary": "Vendor Co", "Reference": "invoice 5"},    # not internal -> ignored
    ]
    # NBF statement: a 500 credit posted (matches deal 500); no 1,400,000 credit
    day.nbf_rows = [
        {"desc": "Inward IPI Payment FT26201AAAAA IPI... AE320960000536060001049", "debit": None, "credit": 500.0},
        {"desc": "Aani Transfer Credit FT26201BBBBB Cr Customer", "debit": None, "credit": 250.0},
    ]

    transfers, s = internal_transfers(day)
    assert s["total"] == 2, s          # the BILL PAYMENT is excluded
    assert s["matched"] == 1, s
    assert s["credit_pending"] == 1, s

    by_ref = {t["Ref"]: t for t in transfers}
    assert by_ref["500"]["Status"] == "Matched", by_ref["500"]
    assert by_ref["500"]["Destination"] == "NBF"
    assert by_ref["500"]["Credit Found In"] == "NBF"
    assert by_ref["1400000"]["Status"] == "Credit pending", by_ref["1400000"]
    assert by_ref["1400000"]["Destination"] == "NBF"
    print("test_transfers: OK")


def test_transfers():
    run()


if __name__ == "__main__":
    run()
