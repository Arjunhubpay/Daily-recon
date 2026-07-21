"""Hubpay's own accounts and the tags used when initiating an internal
movement. A transaction is an internal transfer when its counterparty is one
of these accounts, or when it carries one of the internal tags.

Edit this file to add/adjust accounts or tags — nothing else needs to change.
"""

import re

# IBAN (spaces removed, upper-cased)  ->  (provider, label)
HUBPAY_ACCOUNTS = {
    "AE520960000536010000002": ("Zand", "Zand AED"),
    "AE570960000536010000009": ("Zand", "Zand USD"),
    "AE950960000536010000004": ("Zand", "Zand EUR"),
    "GB75CMFX23016020880620":  ("Corpay", "Corpay (all ccy)"),
    "AE590380000012001794718": ("NBF", "NBF 4718"),
    "AE760380000012001683444": ("NBF", "NBF 3444"),
    "AE070380000012001597289": ("NBF", "NBF 7289"),
    "GB41TCCL04140492650644":  ("Currency Cloud", "Currency Cloud (all ccy)"),
}

# Tags/phrases that mark an internal movement (matched case- and
# spacing-insensitively, so "inter company transfer" == "intercompany transfer"
# and "Internal Transfer 2026-12" matches "internal transfer").
INTERNAL_TAGS = (
    "internal transfer",
    "fis-own account transfer",
    "own account transfer",
    "intercompany transfer",
)

# A movement whose description names one of our own banks as the destination
# is internal even without a tag (e.g. Zand "Transfer to NBF for AR-...").
OWN_BANK_NAMES = ("nbf", "zand", "corpay", "currency cloud")

# Fee sweeps between own accounts are NOT treated as internal transfers here;
# they stay in the normal recon only.
FEE_PHRASES = ("transfer of fee",)

_NON_ALNUM = re.compile(r"[^0-9A-Za-z]")


def norm(s: str) -> str:
    """Normalise an account string for comparison (drop spaces/punctuation)."""
    return _NON_ALNUM.sub("", (s or "")).upper()


# pre-normalised account keys and a digits-only tail index (statements sometimes
# show the domestic account number rather than the full IBAN)
_ACCT_NORM = {norm(k): v for k, v in HUBPAY_ACCOUNTS.items()}


def find_account(*texts):
    """Return (iban, provider, label) if any own account is present in the
    given text fields, else None. Matches the full IBAN, ignoring spacing."""
    blob = norm(" ".join(t for t in texts if t))
    if not blob:
        return None
    for acct, (provider, label) in _ACCT_NORM.items():
        if acct in blob:
            return acct, provider, label
    return None


def has_internal_tag(*texts) -> bool:
    blob = " ".join(t for t in texts if t).lower()
    if any(p in blob for p in FEE_PHRASES):
        return False                      # fee sweeps handled by normal recon only
    despaced = blob.replace(" ", "")
    for tag in INTERNAL_TAGS:
        if tag in blob or tag.replace(" ", "") in despaced:
            return True
    # "transfer to <own bank>" names an own account as the destination
    if "transfer to " in blob and any(b in blob for b in OWN_BANK_NAMES):
        return True
    return False
