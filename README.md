# Recon Daily

Aged daily reconciliation of the internal ledger (`recon-lines`) against the
provider statements — **NBF, Currency Cloud, Corpay, Zand**.

Matching is **two-sided**: an internal line missing from the provider statement
*and* a provider line missing from the internal ledger are both reported.

## Rolling exceptions ledger

There is **one master file, `Open_Exceptions.xlsx`, that is carried forward
and updated every run** — it is always the live list of what is currently
outstanding. Each day:

1. Every currently-open item is **re-checked** against the recent window. If
   the counterparty has since reported it, the item is **cleared** — it moves
   to the `Cleared Log` sheet (stamped with its original *First Seen* date and
   *Days To Clear*) and drops off the open list.
2. Today's **new** breaks (unreconciled even after looking back over the window)
   are **added** as open, dated *First Seen = today*.
3. An item still open **beyond the look-back window (default 10 days)** is
   flagged **`Manual Review`** — a genuine break, not a timing difference.

So: a break you see today automatically disappears from the open list the day
it reconciles; only real, aged breaks remain flagged for manual work.

## Outputs

- **`Output/Open_Exceptions.xlsx`** — the rolling master file. Sheets:
  `Open Exceptions` (live outstanding list) and `Cleared Log` (audit trail).
- **`<date>/Recon_Summary_<date>.xlsx`** — a dated snapshot written into that
  day's folder: recon summary, `Open Exceptions` as of that day, `New Today`
  and `Cleared Today`. Columns follow the existing report schema
  (`Provider, Internal Ref, Provider Ref, Amount, …`).
- **`<date>/dashboard_<date>.html`** and **`Output/dashboard_latest.html`** —
  self-contained dashboards (open / manual-review / cleared-today / matched).

(Output goes to the `Output` sub-folder if it exists, else `_recon_reports`.)

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
python tests/test_engine.py        # matching rules + look-back
python tests/test_ledger.py        # cross-run carry-forward / clearing / manual-review
# or: python -m pytest tests/
```
