#!/usr/bin/env python3
"""
Multi-month reconciliation engine — internal ledger vs provider statements.

Same matching logic as the January recon, run month-by-month:
  - Missing (month M): internal rows dated in M with no provider match anywhere.
  - Overlaps (month M): matched rows whose internal month and provider month
    straddle M's boundary (M-1 <-> M, or M <-> M+1).  These are timing
    differences, not breaks, and they net to zero across the full window.

Providers & keys:
  NBF            -> FT code (in statement description; date encoded as FT+YYDDD)
  Currency Cloud -> Reference, then amount+currency
  Corpay         -> Deal # / Order # (from Supplementary Details or Reference)
  Zand           -> Outgoing Id (from Ref for Account Owner / ref prefix), then amount

Usage:
  python3 recon_engine.py <data_dir> [YYYY-MM YYYY-MM ...]
  - data_dir: folder containing the internal file + all provider files
              (files identified by name/content; provider statements may span
               several months — they are sliced by date in code).
  - months:   optional list of months to report; if omitted, every month present
              in the internal file's Value Date is reported.

Outputs (written next to the data dir, in ./recon_out/):
  summary.csv                 per-month totals per provider
  missing_<YYYY-MM>.csv       internal rows missing in provider, that month
  overlaps_<YYYY-MM>.csv      boundary overlaps, that month
"""
import csv, re, os, sys, datetime, glob

FT = re.compile(r'FT[0-9]{5}[A-Z0-9]{5}')
FEE = {'Debit_FEE_VAT', 'Debit_FEE_REMITTANCE_FEE', 'Credit_FEE_VAT', 'Credit_FEE_REMITTANCE_FEE'}

# ---------- small parsers ----------
def ftx(s):
    m = FT.search((s or '').upper()); return m.group(0) if m else None

def ft_month(ft):
    """FT + YY + DDD (day-of-year) -> (year, month) — NBF's own booking date."""
    m = re.match(r'FT(\d{2})(\d{3})', ft or '')
    if not m: return None
    y = 2000 + int(m.group(1)); doy = int(m.group(2))
    try:
        d = datetime.date(y, 1, 1) + datetime.timedelta(days=doy - 1); return (d.year, d.month)
    except Exception:
        return None

def amtk(v):
    try: return round(abs(float(str(v).replace(',', '').replace('(', '-').replace(')', ''))), 2)
    except Exception: return None

def m_ddmmyyyy(s):
    m = re.match(r'(\d{2})/(\d{2})/(\d{4})', s or '');  return (int(m.group(3)), int(m.group(2))) if m else None

def m_iso(s):
    m = re.match(r'(\d{4})-(\d{2})', s or '');  return (int(m.group(1)), int(m.group(2))) if m else None

def month_add(ym, d):
    y, m = ym; m += d
    while m < 1: m += 12; y -= 1
    while m > 12: m -= 12; y += 1
    return (y, m)

def mlabel(ym):
    return '%04d-%02d' % ym if ym else '?'

# ---------- CSV / file helpers ----------
def read_rows(path):
    with open(path, encoding='utf-8-sig', errors='replace') as fh:
        return list(csv.reader(fh))

def dict_rows(path, skip=0):
    r = read_rows(path)
    if len(r) <= skip: return []
    H = [h.strip() for h in r[skip]]
    return [dict(zip(H, row)) for row in r[skip + 1:]]

def looks_nbf(text):
    head = text[:4000]
    if re.search(r'national bank of fujairah', head, re.I): return True
    if re.search(r'transaction date', head, re.I) and re.search(r'balance', head, re.I): return True
    return len(FT.findall(text[:20000])) >= 3 and bool(re.search(r'description', head, re.I))

def classify(name, sample):
    n = name.lower()
    if 'reconlines' in n or 'recon-lines' in n: return 'INTERNAL'
    if 'transactions_report' in n: return 'CC'
    if 'paymenthistory' in n: return 'CORPAY_PAY'
    if 'settlementreport' in n: return 'CORPAY_SET'
    if 'dealhistory' in n: return 'CORPAY_FX'
    if 'downloaded-statement' in n or 'downloadedstatement' in n: return 'ZAND'
    if 'optransactionhistory' in n or re.search(r'\bnbf\b', n): return 'NBF'
    if n.endswith('.pdf') or n.endswith('.csv'):
        if looks_nbf(sample): return 'NBF'
    return None

def pdf_text(path):
    """Text PDF via pdfplumber; falls back to OCR for scanned PDFs."""
    try:
        import pdfplumber
        out = []
        with pdfplumber.open(path) as pdf:
            for pg in pdf.pages:
                out.append(pg.extract_text() or '')
        t = '\n'.join(out)
        if len(t.strip()) > 40: return t
    except Exception:
        pass
    try:  # scanned -> OCR
        import pypdfium2 as pdfium, pytesseract
        doc = pdfium.PdfDocument(path); out = []
        for i in range(len(doc)):
            out.append(pytesseract.image_to_string(doc[i].render(scale=3.0).to_pil().convert('L')))
        return '\n'.join(out)
    except Exception:
        return ''

def nbf_header_idx(rows):
    for i in range(min(30, len(rows))):
        c = [x.strip().lower() for x in rows[i]]
        if any('description' in x for x in c) and (
           any(x == 'date' or 'transaction date' in x for x in c) or any('debit' in x or 'credit' in x for x in c)):
            return i
    return 0

# ---------- load a data directory ----------
def load(data_dir):
    files = {}
    for p in sorted(glob.glob(os.path.join(data_dir, '*'))):
        if not os.path.isfile(p): continue
        name = os.path.basename(p)
        if name.lower().endswith('.pdf'):
            sample = pdf_text(p); kind = classify(name, sample)
            files.setdefault(kind, []).append(('pdftext', sample, name))
        elif name.lower().endswith('.csv'):
            with open(p, encoding='utf-8-sig', errors='replace') as fh:
                sample = fh.read(6000)
            kind = classify(name, sample)
            files.setdefault(kind, []).append(('csv', p, name))
    return files

def build_provider_index(files):
    """Return dicts used for matching + provider-month lookup."""
    idx = {'nbf_ft': set(), 'cc_ref': set(), 'cc_ac': set(),
           'corpay': set(), 'corpay_set_m': {}, 'corpay_pay_m': {},
           'zand_oid': set(), 'zand_amt': set(), 'zand_m': {}, 'cc_m': {}}

    for kind, entries in files.items():
        for typ, ref, name in entries:
            if kind == 'NBF':
                text = ref if typ == 'pdftext' else open_read(ref)
                if typ == 'csv':
                    rows = read_rows(ref); hi = nbf_header_idx(rows)
                    di = next((i for i, h in enumerate(rows[hi]) if 'description' in h.strip().lower()), 1)
                    for row in rows[hi + 1:]:
                        f = ftx(row[di]) if di < len(row) else None
                        if f: idx['nbf_ft'].add(f)
                else:
                    for f in FT.findall(text.upper()): idx['nbf_ft'].add(f)
            elif kind == 'CC' and typ == 'csv':
                for x in dict_rows(ref):
                    r = (x.get('Reference') or '').strip()
                    mo = m_iso(x.get('Settlement date') or x.get('Completed date') or x.get('Created date'))
                    if r:
                        idx['cc_ref'].add(r)
                        if mo: idx['cc_m'].setdefault(r, mo)
                    idx['cc_ac'].add((amtk(x.get('Amount')), (x.get('Currency') or '').strip()))
            elif kind in ('CORPAY_PAY', 'CORPAY_SET', 'CORPAY_FX') and typ == 'csv':
                # Provider month for overlaps uses SETTLEMENT date (money moved), then
                # payment date. FX deal date is NOT used — a deal struck late in the
                # month but settling T+2 next month is not a boundary break.
                datecol = {'CORPAY_PAY': 'Date', 'CORPAY_SET': 'Settlement Date', 'CORPAY_FX': None}[kind]
                for x in dict_rows(ref):
                    mo = m_iso(x.get(datecol)) if datecol else None
                    for c in ('Deal #', 'Order #', 'Tracker ID', 'Reference'):
                        v = (x.get(c) or '').strip()
                        if v:
                            idx['corpay'].add(v)
                            if mo and kind == 'CORPAY_SET': idx['corpay_set_m'].setdefault(v, mo)
                            if mo and kind == 'CORPAY_PAY': idx['corpay_pay_m'].setdefault(v, mo)
            elif kind == 'ZAND' and typ == 'csv':
                for x in dict_rows(ref):
                    o = (x.get('Outgoing Id') or '').strip()
                    mo = m_iso(x.get('Value date') or x.get('Transaction date'))
                    if o:
                        idx['zand_oid'].add(o)
                        if mo: idx['zand_m'].setdefault(o, mo)
                    a = amtk(x.get('Credit') or x.get('Debit'))
                    if a is not None: idx['zand_amt'].add(a)
    return idx

def open_read(p):
    with open(p, encoding='utf-8-sig', errors='replace') as fh: return fh.read()

# ---------- match one internal row ----------
def deal_of(x):
    sd = (x.get('Supplementary Details') or '').strip(); rf = (x.get('Reference') or '').strip()
    return sd if sd.isdigit() else (rf if rf.isdigit() else '')

def match_row(x, idx):
    """Return (matched: bool, provider_month or None)."""
    p = x.get('Payment Provider')
    if p == 'NBF':
        k = ftx(x.get('Reference'))
        return (bool(k) and k in idx['nbf_ft'], ft_month(k))
    if p == 'CURRENCY_CLOUD':
        r = (x.get('Reference') or '').strip()
        if r in idx['cc_ref']: return (True, idx['cc_m'].get(r))
        if (amtk(x.get('Amount')), (x.get('Currency') or '').strip()) in idx['cc_ac']: return (True, None)
        return (False, None)
    if p == 'CORPAY':
        for cand in [(x.get('Reference') or '').strip(), (x.get('Supplementary Details') or '').strip(),
                     (x.get('Ref for Account Owner') or '').strip(), deal_of(x)]:
            if cand and cand in idx['corpay']:
                return (True, idx['corpay_set_m'].get(cand) or idx['corpay_pay_m'].get(cand))
        return (False, None)
    if p == 'ZAND':
        rf = (x.get('Ref for Account Owner') or '').strip(); pre = (x.get('Reference') or '').split('_')[0]
        if rf in idx['zand_oid']: return (True, idx['zand_m'].get(rf))
        if pre in idx['zand_oid']: return (True, idx['zand_m'].get(pre))
        if amtk(x.get('Amount')) in idx['zand_amt']: return (True, None)
        return (False, None)
    return (False, None)

# ---------- main ----------
def main(data_dir, months=None):
    files = load(data_dir)
    internal_entries = files.get('INTERNAL', [])
    if not internal_entries:
        print('No internal recon-lines file found in', data_dir); return
    internal = dict_rows(internal_entries[0][1])
    idx = build_provider_index(files)

    rows = []
    present = set()
    flagged_by_month = {}   # internal Recon Status != CLEARED -> timing / in-transit / pending
    for x in internal:
        im = m_ddmmyyyy(x.get('Value Date'))
        st = (x.get('Recon Status') or '').strip()
        if st and st != 'CLEARED':
            flagged_by_month.setdefault(im, []).append(x)
        if x.get('Transaction Type') in FEE: continue
        if im: present.add(im)
        matched, pm = match_row(x, idx)
        rows.append((x, im, matched, pm))

    if not months:
        months = sorted(present)
    else:
        months = [tuple(int(a) for a in mm.split('-')) for mm in months]

    outdir = os.path.join(data_dir, 'recon_out'); os.makedirs(outdir, exist_ok=True)
    summary = []
    for M in months:
        prev, nxt = month_add(M, -1), month_add(M, 1)
        provs = {}
        missing, overlaps = [], []
        for x, im, matched, pm in rows:
            p = x.get('Payment Provider')
            if im == M:
                provs.setdefault(p, [0, 0]); provs[p][0] += 1
                if matched: provs[p][1] += 1
                else: missing.append((x, im, pm))
            # overlaps touching M's boundary
            if matched and pm and im and im != pm:
                if (im == M and pm in (prev, nxt)) or (pm == M and im in (prev, nxt)):
                    overlaps.append((x, im, pm))
        # de-dup overlaps recorded once (a row where im==M and pm==nxt is same as counting from M)
        for p in ['CURRENCY_CLOUD', 'NBF', 'ZAND', 'CORPAY']:
            t, mt = provs.get(p, [0, 0])
            summary.append([mlabel(M), p, t, mt, t - mt])
        flagged = flagged_by_month.get(M, [])
        _write(os.path.join(outdir, 'missing_%s.csv' % mlabel(M)), missing, mkey=True)
        _write(os.path.join(outdir, 'overlaps_%s.csv' % mlabel(M)), overlaps, mkey=True, overlap=True)
        _write_flagged(os.path.join(outdir, 'flagged_%s.csv' % mlabel(M)), flagged)
        tot = sum(v[0] for v in provs.values()); mat = sum(v[1] for v in provs.values())
        print('%s  internal=%-6d matched=%-6d missing=%-4d overlaps=%-4d internal-flagged=%-4d' % (
            mlabel(M), tot, mat, tot - mat, len(overlaps), len(flagged)))

    with open(os.path.join(outdir, 'summary.csv'), 'w', newline='') as fh:
        w = csv.writer(fh); w.writerow(['Month', 'Provider', 'Internal', 'Matched', 'Missing']); w.writerows(summary)
    print('\nWrote per-month CSVs + summary.csv to', outdir)

def _write(path, items, mkey=False, overlap=False):
    hdr = ['Provider', 'Internal Month', 'Provider Month', 'Reference', 'Value Date',
           'Amount', 'Currency', 'Transaction Type', 'Description']
    with open(path, 'w', newline='') as fh:
        w = csv.writer(fh); w.writerow(hdr)
        for x, im, pm in items:
            w.writerow([x.get('Payment Provider'), mlabel(im), mlabel(pm) if pm else '',
                        x.get('Reference'), x.get('Value Date'), x.get('Amount'),
                        x.get('Currency'), x.get('Transaction Type'), (x.get('Description') or '')[:90]])

def _write_flagged(path, items):
    """Internal rows the platform's own recon left != CLEARED (timing / in-transit / pending)."""
    hdr = ['Provider', 'Recon Status', 'Value Date', 'Reference', 'Amount', 'Currency',
           'Transaction Type', 'Description']
    with open(path, 'w', newline='') as fh:
        w = csv.writer(fh); w.writerow(hdr)
        for x in items:
            w.writerow([x.get('Payment Provider'), x.get('Recon Status'), x.get('Value Date'),
                        x.get('Reference'), x.get('Amount'), x.get('Currency'),
                        x.get('Transaction Type'), (x.get('Description') or '')[:90]])

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    main(sys.argv[1], sys.argv[2:] or None)
