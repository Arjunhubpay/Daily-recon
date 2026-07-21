# Recon Daily

Aged daily reconciliation of the internal ledger (`recon-lines`) against the
provider statements — **NBF, Currency Cloud, Corpay, Zand**.

Each run does a same-day reconciliation, then **ages** any break against the
previous *N* days of files (default 10). The result is split three ways:

| Bucket | Meaning |
| --- | --- |
| **Matched** | Reconciled against the **same day's** data. |
| **Cleared** | Was a same-day break, but reconciles against a **prior day** within the look-back window — a timing difference. Stamped with the earlier date it matched against. |
| **Exception** | Not present on **either side** across the run date *and* the whole look-back window. **Needs manual checking.** |

Matching is **two-sided**: an internal line missing from the provider statement
*and* a provider line missing from the internal ledger are both reported.

Artefacts written per run (to the `Output` sub-folder if one exists, else
`_recon_reports`):
- `Recon_Results.xlsx` — latest run, columns matching the existing report
  (`Run Date, Provider, Internal Ref, Provider Ref, Value Date,
  Amount (Internal), Amount (Provider), Difference, …`) across
  Summary / Exceptions / Cleared / Matched sheets.
- `dashboard_latest.html` — self-contained dashboard for the latest run.
- `<date>/recon_report_<date>.xlsx` and `<date>/dashboard_<date>.html` — dated archive copies.

`Difference` compares **magnitudes** (providers store signed amounts, the
internal ledger stores magnitudes), so a matched pair with equal value shows
`0.00`; a non-zero value is a genuine amount discrepancy worth reviewing.

---

## Folder layout (important)

Put **each day's files in a sub-folder named by date** (`YYYY-MM-DD`). That is
how the tool "checks by date" and finds the look-back history:

```
<share folder>/
├── 2026-07-21/                     <- today's upload
│   ├── recon-lines-....csv
│   ├── transactions_report-....csv
│   ├── PaymentHistory-....xlsx
│   ├── SettlementReport-....xlsx
│   ├── DealHistory-....xlsx
│   ├── OpTransactionHistory-....xlsx
│   └── downloaded-statement-....csv
├── 2026-07-20/                     <- yesterday
│   └── ...
└── _recon_reports/                 <- outputs land here (auto-created)
    ├── 2026-07-21/
    │   ├── recon_report_2026-07-21.xlsx
    │   └── dashboard_2026-07-21.html
    └── dashboard_latest.html
```

File **names** are matched by keyword (`recon-lines`, `transactions_report`,
`PaymentHistory`, `SettlementReport`, `DealHistory`, `OpTransactionHistory`,
`downloaded-statement`) — the rest of the name and the date suffix don't matter.
Both `.csv` and `.xlsx` work.

> **SharePoint:** the simplest setup is to *sync* the SharePoint library to your
> machine (via the OneDrive client). It then appears as a normal local folder,
> and you point `--root` at it — no API keys needed.

---

## Install

Requires **Python 3.9+**.

```bash
pip install -r requirements.txt
```

## Run

```bash
python -m recon --root "/path/to/your/share/folder"
```

Options:

| Flag | Default | Description |
| --- | --- | --- |
| `--root` | *(required)* | Share folder containing the dated sub-folders. |
| `--date` | today | Run date, `YYYY-MM-DD`. |
| `--lookback` | `10` | Days to look back when aging a break. |
| `--out` | `<root>/_recon_reports` | Output folder. |
| `--quiet` | off | Suppress the console summary. |

---

## Schedule it daily

### Windows (Task Scheduler)
1. Edit **`run_recon.bat`** and set `RECON_ROOT` to your share folder.
2. Task Scheduler → *Create Basic Task* → trigger **Daily** → action **Start a
   program** → point it at `run_recon.bat`.

### macOS / Linux (cron)
1. Edit **`run_recon.sh`** and set `RECON_ROOT`; `chmod +x run_recon.sh`.
2. `crontab -e`, then (e.g. every day at 07:30):
   ```
   30 7 * * * /path/to/Daily-recon/run_recon.sh >> /path/to/recon.log 2>&1
   ```

---

## Matching rules

- **NBF** — the `FT#####XXXXX` code embedded in the statement description
  (lines containing "computer generated" are ignored).
- **Currency Cloud** — `Reference`, then an amount + currency fallback.
- **Corpay** — `Deal #` taken from the internal `Supplementary Details` (else
  `Reference`), matched across **PaymentHistory**, **SettlementReport** (paid
  only) and **DealHistory**.
- **Zand** — `Outgoing Id` (the internal `Reference` with a trailing `_UUID`
  stripped), then `Ref for Account Owner`, then an amount fallback.

Fee rows (`Debit_FEE_VAT`, `Debit_FEE_REMITTANCE_FEE`) are excluded.

## Test

```bash
python tests/test_engine.py        # or: python -m pytest tests/
```
