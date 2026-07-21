"""
Recon Daily — reconciliation engine.

Reconciles the internal ledger (recon-lines) against provider statements
(NBF, Currency Cloud, Corpay, Zand) for a given run date, then ages any
same-day break against the previous N days of files:

  * matched   — reconciled against the SAME day's provider/internal data
  * cleared   — was a same-day break, but reconciles against a PRIOR day's
                data (a timing difference); stamped with that earlier date
  * exception — not present on either side across the run date AND the
                look-back window; needs manual checking

Matching is two-sided: an internal line missing from the provider statement
and a provider line missing from the internal ledger are both reported.

The matching rules are ported 1:1 from the browser tool:
  NBF            : FT code embedded in the statement description
  Currency Cloud : Reference, then amount+currency fallback
  Corpay         : Deal # (from Supplementary Details, else Reference) across
                   Payments / Settlements / FX deals
  Zand           : Outgoing Id (Reference with trailing _UUID stripped),
                   then Ref for Account Owner, then amount fallback
"""

from __future__ import annotations

import csv
import datetime as _dt
import re
from dataclasses import dataclass, field
from pathlib import Path

# ---------------------------------------------------------------------------
# File-type detection (same keywords as the browser tool)
# ---------------------------------------------------------------------------
KEYWORDS = [
    ("recon-lines", "Internal"),
    ("transactions_report", "Currency Cloud"),
    ("PaymentHistory", "Corpay Payments"),
    ("SettlementReport", "Corpay Settlements"),
    ("DealHistory", "Corpay FX"),
    ("OpTransactionHistory", "NBF"),
    ("downloaded-statement", "Zand"),
]

FEE_TYPES = {"Debit_FEE_VAT", "Debit_FEE_REMITTANCE_FEE"}
PROVIDERS = ["NBF", "Currency Cloud", "Corpay", "Zand"]
# internal 'Payment Provider' code -> display name
INTERNAL_PROVIDER_CODE = {
    "NBF": "NBF",
    "CURRENCY_CLOUD": "Currency Cloud",
    "CORPAY": "Corpay",
    "ZAND": "Zand",
}

FT_RE = re.compile(r"FT[0-9]{5}[A-Z0-9]{5}")
_ZAND_UUID_RE = re.compile(r"_[0-9a-fA-F]{8}$")


# ---------------------------------------------------------------------------
# Value helpers (ported from the verified JS)
# ---------------------------------------------------------------------------
def num(x):
    """Parse a possibly-messy numeric string. Returns float or None."""
    if x is None:
        return None
    s = str(x).strip().replace(",", "")
    if s == "" or s.lower() == "nan":
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^0-9.\-]", "", s)
    if s in ("", "-", "."):
        return None
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def amt_key(v):
    """Round to cents for comparison. Returns rounded float or None."""
    return None if v is None else round(v * 100) / 100


def extract_ft(text):
    if not text:
        return None
    m = FT_RE.search(str(text))
    return m.group(0) if m else None


def strip_zand(ref):
    if ref is None:
        return ""
    return _ZAND_UUID_RE.sub("", str(ref).strip())


def detect(name: str):
    low = name.lower()
    for kw, kind in KEYWORDS:
        if kw.lower() in low:
            return kind
    return None


# ---------------------------------------------------------------------------
# File reading  ->  list[dict] with the same semantics as JS toObjects()
# ---------------------------------------------------------------------------
def _rows_from_csv(path: Path):
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        return [[("" if c is None else str(c)) for c in row]
                for row in csv.reader(fh)]


def _rows_from_xlsx(path: Path):
    import openpyxl  # local import so CSV-only setups need no dependency

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(["" if c is None else _cell_str(c) for c in row])
    wb.close()
    return rows


def _cell_str(c):
    if isinstance(c, _dt.datetime):
        # match a spreadsheet-style date; drop midnight time component
        if c.hour == c.minute == c.second == 0:
            return c.strftime("%Y-%m-%d")
        return c.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(c, _dt.date):
        return c.strftime("%Y-%m-%d")
    if isinstance(c, float) and c.is_integer():
        return str(int(c))
    return str(c)


def read_rows(path: Path):
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xls"):
        return _rows_from_xlsx(path)
    return _rows_from_csv(path)


def to_objects(rows, skip=0):
    """Rows (list[list[str]]) -> list[dict] using row `skip` as the header."""
    if len(rows) <= skip:
        return []
    header = [h.strip() for h in rows[skip]]
    out = []
    for r in range(skip + 1, len(rows)):
        cells = rows[r]
        if len(cells) == 1 and cells[0].strip() == "":
            continue
        obj = {}
        for i, h in enumerate(header):
            obj[h] = (cells[i].strip() if i < len(cells) else "")
        out.append(obj)
    return out


# ---------------------------------------------------------------------------
# One day's worth of files, parsed into the structures matching needs
# ---------------------------------------------------------------------------
@dataclass
class DayData:
    date: _dt.date
    internal: dict = field(default_factory=dict)   # provider display name -> [rows]
    cc_by_ref: dict = field(default_factory=dict)   # ref -> [cc rows]
    cc_rows: list = field(default_factory=list)      # all cc rows (for amount fallback)
    cp_pay: dict = field(default_factory=dict)
    cp_setl: dict = field(default_factory=dict)
    cp_fx: dict = field(default_factory=dict)
    nbf_idx: dict = field(default_factory=dict)      # FT -> count
    zand_idx: dict = field(default_factory=dict)     # outgoing id -> [{amt}]
    files: list = field(default_factory=list)        # (name, kind) seen
    # raw statement rows kept for internal-transfer detection / credit checks
    nbf_rows: list = field(default_factory=list)     # {desc, debit, credit}
    zand_rows: list = field(default_factory=list)    # {oid, details, credit, debit}
    corpay_pay_rows: list = field(default_factory=list)  # full PaymentHistory rows

    @property
    def cp_all(self):
        return set(self.cp_pay) | set(self.cp_setl) | set(self.cp_fx)


def load_day(date: _dt.date, file_paths) -> DayData:
    """file_paths: iterable of pathlib.Path for that day's files."""
    day = DayData(date=date)

    def pick(kw):
        return [p for p in file_paths if kw.lower() in p.name.lower()]

    # internal
    internal_raw = []
    for p in pick("recon-lines"):
        internal_raw.extend(to_objects(read_rows(p), 0))
    internal = [r for r in internal_raw if r.get("Transaction Type") not in FEE_TYPES]
    for name in PROVIDERS:
        day.internal[name] = []
    for r in internal:
        disp = INTERNAL_PROVIDER_CODE.get((r.get("Payment Provider") or "").strip())
        if disp:
            day.internal[disp].append(r)

    # currency cloud
    for p in pick("transactions_report"):
        for r in to_objects(read_rows(p), 0):
            k = (r.get("Reference") or "").strip()
            day.cc_by_ref.setdefault(k, []).append(r)
            day.cc_rows.append(r)

    # corpay
    def corpay_index(kw, amt_col, paid_only):
        idx = {}
        for p in pick(kw):
            for r in to_objects(read_rows(p), 0):
                if paid_only and (r.get("Status") or "").strip().lower() != "paid":
                    continue
                deal = (r.get("Deal #") or "").strip()
                if deal:
                    idx.setdefault(deal, []).append({"amt": num(r.get(amt_col))})
        return idx

    day.cp_pay = corpay_index("PaymentHistory", "Amount", False)
    day.cp_setl = corpay_index("SettlementReport", "Settlement Amount", True)
    day.cp_fx = corpay_index("DealHistory", "Purchased", False)
    for p in pick("PaymentHistory"):
        day.corpay_pay_rows.extend(to_objects(read_rows(p), 0))

    # nbf — header on row index 7; find the Description column
    for p in pick("OpTransactionHistory"):
        rows = to_objects(read_rows(p), 7)
        if not rows:
            continue
        desc_key = next((k for k in rows[0].keys() if "Description" in k), "")
        dr_key = next((k for k in rows[0].keys() if "Debit" in k), "")
        cr_key = next((k for k in rows[0].keys() if "Credit" in k), "")
        for r in rows:
            desc = r.get(desc_key, "") if desc_key else ""
            if "computer generated" in desc.lower():
                continue
            day.nbf_rows.append({"desc": desc, "debit": num(r.get(dr_key)),
                                 "credit": num(r.get(cr_key))})
            ft = extract_ft(desc)
            if ft:
                day.nbf_idx[ft] = day.nbf_idx.get(ft, 0) + 1

    # zand
    for p in pick("downloaded-statement"):
        for r in to_objects(read_rows(p), 0):
            details = next((r[k] for k in r if k.startswith("Details")), "")
            day.zand_rows.append({"oid": (r.get("Outgoing Id") or "").strip(),
                                  "details": details,
                                  "credit": num(r.get("Credit")), "debit": num(r.get("Debit"))})
            oid = (r.get("Outgoing Id") or "").strip()
            if not oid:
                continue
            amt = num(r.get("Credit"))
            if amt is None:
                amt = num(r.get("Debit"))
            day.zand_idx.setdefault(oid, []).append({"amt": amt})

    for p in file_paths:
        day.files.append((p.name, detect(p.name)))
    return day


# ---------------------------------------------------------------------------
# Matching — internal line  ->  a given day's provider index
# Returns (matched: bool, provider_ref: str, method: str, provider_amount)
# provider_amount is the amount recorded on the provider side (None if the
# provider index carries no amount, e.g. NBF), used for the Difference column.
# ---------------------------------------------------------------------------
def match_internal(provider: str, ir: dict, day: DayData):
    if provider == "NBF":
        ref = (ir.get("Reference") or "").strip()
        if day.nbf_idx.get(ref):
            return True, ref, "FT code", None
        return False, ref, "", None

    if provider == "Currency Cloud":
        ref = (ir.get("Reference") or "").strip()
        if day.cc_by_ref.get(ref):
            return True, ref, "Reference", amt_key(num(day.cc_by_ref[ref][0].get("Amount")))
        amt = amt_key(num(ir.get("Amount")))
        cur = ir.get("Currency")
        for r in day.cc_rows:
            if amt_key(num(r.get("Amount"))) == amt and r.get("Currency") == cur:
                return True, (r.get("Reference") or ""), "Amount+Currency", amt_key(num(r.get("Amount")))
        return False, ref, "", None

    if provider == "Corpay":
        sd = (ir.get("Supplementary Details") or "").strip()
        ref = (ir.get("Reference") or "").strip()
        deal = sd if re.fullmatch(r"\d+", sd) else (ref if re.fullmatch(r"\d+", ref) else "")
        if deal and deal in day.cp_all:
            src = (day.cp_pay.get(deal) or day.cp_setl.get(deal) or day.cp_fx.get(deal) or [{}])
            where = ("Payments" if deal in day.cp_pay
                     else "Settlements" if deal in day.cp_setl else "FX Deals")
            return True, deal, f"Deal # ({where})", amt_key(src[0].get("amt"))
        return False, (deal or ref), "", None

    if provider == "Zand":
        clean = strip_zand(ir.get("Reference"))
        rfao = (ir.get("Ref for Account Owner") or "").strip()
        if day.zand_idx.get(clean):
            return True, clean, "Outgoing Id", amt_key(day.zand_idx[clean][0]["amt"])
        if day.zand_idx.get(rfao):
            return True, rfao, "Ref for Account Owner", amt_key(day.zand_idx[rfao][0]["amt"])
        amt = amt_key(num(ir.get("Amount")))
        for k, lst in day.zand_idx.items():
            for r in lst:
                if amt_key(r["amt"]) == amt:
                    return True, k, "Amount", amt_key(r["amt"])
        return False, clean, "", None

    return False, "", "", None


# ---------------------------------------------------------------------------
# Provider-side rows (for the reverse direction: provider not in internal)
# Each entry: {"key", "amt", "cur", "extra"} keyed for reverse matching.
# ---------------------------------------------------------------------------
def provider_rows(provider: str, day: DayData):
    rows = []
    if provider == "NBF":
        for ft, cnt in day.nbf_idx.items():
            for _ in range(cnt):
                rows.append({"key": ft, "amt": None, "cur": ""})
    elif provider == "Currency Cloud":
        for r in day.cc_rows:
            rows.append({"key": (r.get("Reference") or "").strip(),
                         "amt": amt_key(num(r.get("Amount"))),
                         "cur": r.get("Currency") or ""})
    elif provider == "Corpay":
        for deal in day.cp_all:
            src = (day.cp_pay.get(deal) or day.cp_setl.get(deal) or day.cp_fx.get(deal) or [{}])
            rows.append({"key": deal, "amt": amt_key(src[0].get("amt")), "cur": ""})
    elif provider == "Zand":
        for oid, lst in day.zand_idx.items():
            for r in lst:
                rows.append({"key": oid, "amt": amt_key(r["amt"]), "cur": ""})
    return rows


def internal_match_keys(provider: str, day: DayData):
    """The set of keys (and amounts) that internal lines would match on,
    used to decide whether a provider row has an internal counterpart."""
    keys = set()
    amts = set()
    for ir in day.internal.get(provider, []):
        if provider == "NBF":
            keys.add((ir.get("Reference") or "").strip())
        elif provider == "Currency Cloud":
            keys.add((ir.get("Reference") or "").strip())
            a = amt_key(num(ir.get("Amount")))
            if a is not None:
                amts.add((a, ir.get("Currency")))
        elif provider == "Corpay":
            sd = (ir.get("Supplementary Details") or "").strip()
            ref = (ir.get("Reference") or "").strip()
            deal = sd if re.fullmatch(r"\d+", sd) else (ref if re.fullmatch(r"\d+", ref) else "")
            if deal:
                keys.add(deal)
        elif provider == "Zand":
            keys.add(strip_zand(ir.get("Reference")))
            keys.add((ir.get("Ref for Account Owner") or "").strip())
            a = amt_key(num(ir.get("Amount")))
            if a is not None:
                amts.add((a, ""))
    keys.discard("")
    return keys, amts


def provider_row_has_internal(provider: str, prow: dict, day: DayData):
    keys, amts = internal_match_keys(provider, day)
    if prow["key"] and prow["key"] in keys:
        return True
    if provider in ("Currency Cloud", "Zand") and prow["amt"] is not None:
        cur = prow["cur"] if provider == "Currency Cloud" else ""
        if (prow["amt"], cur) in amts:
            return True
    return False


def open_item_reconciles(provider, side, key, amount, currency, day: DayData):
    """Re-test a carried-forward open exception against a later day's data:
    has the counterparty since reported it? `key` is the item's matching key
    (provider_ref recorded when the exception was opened)."""
    key = (key or "").strip()
    a = amt_key(num(amount))
    if side == "internal":
        # has the PROVIDER now reported this internal item?
        if provider == "NBF":
            return bool(day.nbf_idx.get(key))
        if provider == "Currency Cloud":
            if day.cc_by_ref.get(key):
                return True
            return any(amt_key(num(r.get("Amount"))) == a and r.get("Currency") == currency
                       for r in day.cc_rows)
        if provider == "Corpay":
            return key in day.cp_all
        if provider == "Zand":
            if day.zand_idx.get(key):
                return True
            return any(amt_key(r["amt"]) == a
                       for lst in day.zand_idx.values() for r in lst)
        return False
    else:
        # provider-side: has the INTERNAL ledger now recorded this item?
        keys, amts = internal_match_keys(provider, day)
        if key and key in keys:
            return True
        if provider in ("Currency Cloud", "Zand") and a is not None:
            cur = currency if provider == "Currency Cloud" else ""
            return (a, cur) in amts
        return False


# ---------------------------------------------------------------------------
# Internal transfers (intra-Hubpay: money moving between two own accounts).
# Detected from the PROVIDER STATEMENTS, because a leg is often not tagged in
# the internal ledger (e.g. a Corpay "own account transfer" only shows its
# INTERCOMPANY nature in the Corpay file). For each OUTGOING internal leg we
# verify the CREDIT actually posted in the destination bank's statement; if it
# has not, the transfer is flagged "Credit pending".
# ---------------------------------------------------------------------------
def _dest_provider(*texts):
    from .accounts import find_account
    acct = find_account(*texts)
    if acct:
        return acct[1]                       # provider owning the beneficiary account
    t = " ".join(x for x in texts if x).lower()
    if "nbf" in t or "national bank of fujairah" in t:
        return "NBF"
    if "currency cloud" in t or "gbp cc" in t or "- cc" in t or " cc " in t:
        return "Currency Cloud"
    if "corpay" in t or "cmfx" in t:
        return "Corpay"
    if "zand" in t:
        return "Zand"
    return ""


def _build_credit_pools(day: DayData):
    """All incoming amounts per provider (regardless of whether the credit is
    itself labelled internal) — used to confirm a transfer's credit posted."""
    import collections
    pools = collections.defaultdict(collections.Counter)
    for r in day.nbf_rows:
        if r["credit"] is not None:
            pools["NBF"][amt_key(r["credit"])] += 1
    for r in day.zand_rows:
        if r["credit"] is not None:
            pools["Zand"][amt_key(r["credit"])] += 1
    for r in day.cc_rows:
        a = amt_key(num(r.get("Amount")))
        if a is not None:
            pools["Currency Cloud"][abs(a)] += 1
    for lst in day.cp_setl.values():
        for x in lst:
            if x.get("amt") is not None:
                pools["Corpay"][amt_key(x["amt"])] += 1
    return pools


def internal_transfers(day: DayData):
    """Return (transfers, summary). Outgoing internal legs are gathered from the
    provider statements and each is confirmed against the destination bank's
    credits; an unconfirmed one is 'Credit pending'."""
    from .accounts import has_internal_tag, find_account, OWN_BANK_NAMES, FEE_PHRASES

    def is_fee(t):
        return any(p in t.lower() for p in FEE_PHRASES)

    legs = []   # outgoing internal-transfer legs

    # Corpay — payments whose beneficiary is an own account / intercompany
    for r in day.corpay_pay_rows:
        purpose = (r.get("Purpose Of Payment") or "").strip().upper()
        ident, ben, ref = r.get("Identifier", ""), r.get("Beneficiary", ""), r.get("Reference", "")
        internal = (purpose == "INTERCOMPANY PAYMENT" or has_internal_tag(ident, ben, ref)
                    or find_account(ident, ben, ref) is not None)
        if internal:
            legs.append({"src": "Corpay", "dest": _dest_provider(ident, ben, ref),
                         "amt": amt_key(num(r.get("Amount"))), "cur": (r.get("Currency") or ""),
                         "ref": (r.get("Deal #") or "").strip(), "note": (ident or ref)[:55]})

    # Zand — debit rows (money leaving) that are internal movements
    for r in day.zand_rows:
        det = r["details"]
        if is_fee(det) or not r["debit"]:
            continue
        low = det.lower()
        internal = (has_internal_tag(det) or find_account(det) is not None
                    or ("transfer to " in low and any(b in low for b in OWN_BANK_NAMES)))
        if internal:
            legs.append({"src": "Zand", "dest": _dest_provider(det), "amt": amt_key(r["debit"]),
                         "cur": "AED", "ref": r["oid"], "note": det[:55]})

    # NBF — debit rows (money leaving) that are internal movements
    for r in day.nbf_rows:
        det = r["desc"]
        if not r["debit"]:
            continue
        if "internal transf" in det.lower() or has_internal_tag(det) or find_account(det) is not None:
            legs.append({"src": "NBF", "dest": _dest_provider(det), "amt": amt_key(r["debit"]),
                         "cur": "AED", "ref": extract_ft(det) or "", "note": det[:55]})

    # Currency Cloud — payments whose beneficiary is an own account
    for r in day.cc_rows:
        acct = find_account(r.get("Beneficiary IBAN", ""), r.get("Beneficiary account number", ""))
        if acct:
            a = amt_key(num(r.get("Amount")))
            legs.append({"src": "Currency Cloud", "dest": acct[1],
                         "amt": (abs(a) if a is not None else None),
                         "cur": (r.get("Currency") or ""), "ref": (r.get("Reference") or "").strip(),
                         "note": "beneficiary = own account"})

    # confirm each outgoing leg against the destination bank's credits
    pools = _build_credit_pools(day)
    transfers = []
    for L in legs:
        amt, dest = L["amt"], L["dest"]
        candidates = [dest] if dest else list(pools.keys())
        matched_in = None
        for prov in candidates:
            if prov and amt is not None and pools[prov].get(amt, 0) > 0:
                pools[prov][amt] -= 1
                matched_in = prov
                break
        transfers.append({
            "Status": "Matched" if matched_in else "Credit pending",
            "Amount": _fmt_amt(amt), "Currency": L["cur"],
            "Source": L["src"], "Destination": dest or "?",
            "Credit Found In": matched_in or "",
            "Ref": L["ref"], "Note": L["note"],
        })

    order = {"Credit pending": 0, "Matched": 1}
    transfers.sort(key=lambda t: (order.get(t["Status"], 9), t["Source"], t["Ref"]))
    summary = {
        "total": len(transfers),
        "matched": sum(1 for t in transfers if t["Status"] == "Matched"),
        "credit_pending": sum(1 for t in transfers if t["Status"] == "Credit pending"),
        "debit_pending": 0, "amount_mismatch": 0, "unpaired": 0,
    }
    return transfers, summary


def reconciles_in_window(provider, side, key, amount, currency, day_map, dates):
    """True if the open item reconciles in ANY of the given days (tolerates
    skipped runs, e.g. weekends). Returns the first matching date, or None."""
    for d in dates:
        if d in day_map and open_item_reconciles(provider, side, key, amount, currency, day_map[d]):
            return d
    return None


# ---------------------------------------------------------------------------
# Result records
# ---------------------------------------------------------------------------
def _fmt_amt(v):
    """Format an amount for output; '' for None, trims trailing .0 noise."""
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return f"{v:.2f}"
    return str(v)


@dataclass
class Record:
    provider: str
    side: str            # "internal" or "provider"
    status: str          # "matched" | "cleared" | "exception"
    internal_ref: str = ""
    provider_ref: str = ""
    value_date: str = ""
    amount_internal: object = None
    amount_provider: object = None
    currency: str = ""
    txn_type: str = ""
    method: str = ""
    cleared_date: str = ""   # the prior date it reconciled against
    description: str = ""

    @property
    def difference(self):
        # Compare magnitudes: providers use signed amounts (debit negative)
        # while the internal ledger stores magnitudes, so a sign mismatch is a
        # convention difference, not a real discrepancy.
        if isinstance(self.amount_internal, (int, float)) and \
           isinstance(self.amount_provider, (int, float)):
            return round(abs(self.amount_internal) - abs(self.amount_provider), 2)
        return None

    def as_dict(self):
        # Leading columns mirror the existing Recon_Results.xlsx schema.
        return {
            "Provider": self.provider,
            "Internal Ref": self.internal_ref,
            "Provider Ref": self.provider_ref,
            "Value Date": self.value_date,
            "Amount (Internal)": _fmt_amt(self.amount_internal),
            "Amount (Provider)": _fmt_amt(self.amount_provider),
            "Difference": _fmt_amt(self.difference),
            "Currency": self.currency,
            "Side": self.side,
            "Status": self.status,
            "Type": self.txn_type,
            "Method": self.method,
            "Cleared Against": self.cleared_date,
            "Description": self.description,
        }


# ---------------------------------------------------------------------------
# Full reconciliation with look-back aging
# ---------------------------------------------------------------------------
def reconcile(run_date: _dt.date, day_map: dict, lookback_days: int = 10,
              two_sided: bool = False):
    """
    day_map: {date -> DayData}. Must contain run_date; may contain prior days.
    two_sided: also flag provider rows missing from the internal ledger
        (Missing in Internal). Default False — the recon is internal-driven,
        matching the existing process. Enable for a full bidirectional check.
    Returns a result dict with matched / cleared / exceptions / per-provider.
    """
    today = day_map[run_date]
    prior_dates = sorted(
        [d for d in day_map
         if run_date - _dt.timedelta(days=lookback_days) <= d < run_date],
        reverse=True,
    )

    matched, cleared, exceptions = [], [], []

    for provider in PROVIDERS:
        # ---- internal side ----------------------------------------------
        for ir in today.internal.get(provider, []):
            iamt = amt_key(num(ir.get("Amount")))
            ok, pref, method, pamt = match_internal(provider, ir, today)
            if ok:
                matched.append(Record(
                    provider, "internal", "matched",
                    internal_ref=(ir.get("Reference") or ""),
                    provider_ref=pref,
                    value_date=(ir.get("Value Date") or ""),
                    amount_internal=iamt, amount_provider=pamt,
                    currency=(ir.get("Currency") or ""),
                    txn_type=(ir.get("Transaction Type") or ""),
                    method=method))
                continue
            # aged look-back
            cleared_hit = None
            for pd in prior_dates:
                ok2, pref2, method2, pamt2 = match_internal(provider, ir, day_map[pd])
                if ok2:
                    cleared_hit = (pd, pref2, method2, pamt2)
                    break
            base = dict(
                internal_ref=(ir.get("Reference") or ""),
                value_date=(ir.get("Value Date") or ""),
                amount_internal=iamt,
                currency=(ir.get("Currency") or ""),
                txn_type=(ir.get("Transaction Type") or ""))
            if cleared_hit:
                pd, pref2, method2, pamt2 = cleared_hit
                cleared.append(Record(
                    provider, "internal", "cleared", provider_ref=pref2,
                    amount_provider=pamt2,
                    method=method2, cleared_date=pd.isoformat(),
                    description=f"Not in {provider} on {run_date}; reconciled against {pd}",
                    **base))
            else:
                shown = base["internal_ref"] or pref
                exceptions.append(Record(
                    provider, "internal", "exception", provider_ref=pref,
                    method="Missing in Provider",
                    description=f"Ref {shown} not in {provider} on "
                                f"{run_date} or prior {lookback_days} days",
                    **base))

        # ---- provider side (missing in internal) — only when two-sided ---
        for prow in (provider_rows(provider, today) if two_sided else []):
            if provider_row_has_internal(provider, prow, today):
                continue  # already covered by an internal match
            cleared_hit = None
            for pd in prior_dates:
                if provider_row_has_internal(provider, prow, day_map[pd]):
                    cleared_hit = pd
                    break
            if cleared_hit:
                cleared.append(Record(
                    provider, "provider", "cleared", provider_ref=prow["key"],
                    amount_provider=prow["amt"], currency=prow["cur"],
                    method="Provider→Internal", cleared_date=cleared_hit.isoformat(),
                    description=f"{provider} {prow['key']} not in internal on "
                                f"{run_date}; reconciled against {cleared_hit}"))
            else:
                exceptions.append(Record(
                    provider, "provider", "exception", provider_ref=prow["key"],
                    amount_provider=prow["amt"], currency=prow["cur"],
                    method="Missing in Internal",
                    description=f"{provider} {prow['key']} not in internal on "
                                f"{run_date} or prior {lookback_days} days"))

    per_provider = []
    for p in PROVIDERS:
        per_provider.append({
            "provider": p,
            "matched": sum(1 for r in matched if r.provider == p),
            "cleared": sum(1 for r in cleared if r.provider == p),
            "exceptions": sum(1 for r in exceptions if r.provider == p),
        })

    internal_count = sum(len(v) for v in today.internal.values())
    return {
        "run_date": run_date.isoformat(),
        "lookback_days": lookback_days,
        "prior_dates": [d.isoformat() for d in prior_dates],
        "matched": matched,
        "cleared": cleared,
        "exceptions": exceptions,
        "per_provider": per_provider,
        "internal_count": internal_count,
    }
