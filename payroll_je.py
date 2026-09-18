#!/usr/bin/env python3
"""Convert a "DRAFT by-department" payroll workbook (.xlsx) into a QuickBooks
Online journal-entry import CSV.

Usage:
    python payroll_je.py <input.xlsx> --journal-no JJ1234 --date D/M/YYYY --output <out.csv>

Deterministic parsing only: stdlib + openpyxl, single pass over the sheet.
Unknown account names / classes / non-numeric amounts are never guessed: the row
is written with a "<<REVIEW: ...>>" placeholder, a warning goes to stderr, and the
script exits non-zero (the CSV is still written).
"""

import argparse
import csv
import re
import sys
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from openpyxl import load_workbook

# Single source of truth for QuickBooks account names. Exact match after strip().
ACCOUNT_NAME_MAP = {
    "6050 Wages & Salaries": "6050 Wages & Salaries -",
    "6046 CPP Expense": "6046 CPP Expense -",
    "6047 EI Expense": "6047 EI Expense -",
    "6049 Vacation expense": "6049 Vacation expense",
    "6044 Tips & Gratuities": "6044 Tips & Gratuities",
    "2024 Staff Deposits": "2024 Staff Deposits",
    "6043 Benefits": "6043 Employee Health Benefits",
    "6043 Employer Health Benefits": "6043 Employee Health Benefits",
    "6043 Employee Health Benefits": "6043 Employee Health Benefits",
    "6007 Utilities": "6007 Utilities (DS)",
    "6007 Utilities (DS)": "6007 Utilities (DS)",
    "2016 Wages Payable": "2016 Wages Payable",
    "2017 Federal Income Tax Payable": "2017 Federal Income Tax Payable",
    "2014 CPP Payable": "2014 CPP Payable",
    "2013 EI Payable": "2013 EI Payable",
    "2015 RRSP Payable": "2015 RRSP Payable",
}

CLASS_NAME_FIXUPS = {
    "0097-INFORMATION TECHNOLOGY": "0097-INFORMATION TECHNOLOGY (IT)",
}

CSV_COLUMNS = ['*JournalNo', '*JournalDate', 'Memo', '*AccountName', 'Debits', 'Credits',
               'Description', 'Name', 'Location', 'Class']

NOISE_PREFIXES = ('company-wide', 'notes', 'draft', 'payroll journal', 'check')
NUMBERED_NOTE = re.compile(r'^\d+\.')
DEPT_HEADER = re.compile(r'^(\S+)\s+[—–-]\s+(.+)$')
CLASS_PATTERN = re.compile(r'^\d{4}-\S.*$')
CENT = Decimal('0.01')


def review(text):
    return f'<<REVIEW: {text}>>'


def parse_date(s):
    """Strict D/M/YYYY. Returns (iso_for_qb, human) or raises ValueError."""
    m = re.fullmatch(r'\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*', s or '')
    if not m:
        raise ValueError(f"--date must be D/M/YYYY (e.g. 1/8/2026), got {s!r}")
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if mo > 12:
        raise ValueError(f"--date {s!r}: month {mo} > 12 — this looks like MM/DD/YYYY. Use D/M/YYYY.")
    try:
        dt = date(y, mo, d)
    except ValueError as e:
        raise ValueError(f"--date {s!r} is not a valid calendar date: {e}")
    return dt.strftime('%d/%m/%Y'), dt.strftime('%d %B %Y')


def to_amount(v):
    """Returns Decimal or None for blank. Raises InvalidOperation for junk."""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip().replace('$', '').replace(',', '')
        if s == '':
            return None
        if s.startswith('='):
            raise InvalidOperation(v)   # un-evaluated formula (no cached value)
        return Decimal(s).quantize(CENT, rounding=ROUND_HALF_UP)
    if isinstance(v, bool):
        raise InvalidOperation(v)
    return Decimal(str(v)).quantize(CENT, rounding=ROUND_HALF_UP)


def normalize_class(header_text, row_no, warnings):
    """Extract class name from a department header; placeholder if unparseable."""
    m = DEPT_HEADER.match(header_text)
    raw = m.group(2).strip() if m else ''
    cut = raw.find(' (')
    if cut != -1:
        raw = raw[:cut].strip()
    raw = CLASS_NAME_FIXUPS.get(raw, raw)
    if not raw or not CLASS_PATTERN.match(raw):
        warnings.append(f"WARNING: unparseable class in header '{header_text}' at row {row_no} — using placeholder, fix before import")
        return review(header_text), True
    return raw, False


def fmt(d):
    return '' if d is None else f'{d:.2f}'


def convert(xlsx_path, journal_no, date_str):
    """Returns (rows, warnings, flagged_count, debits_total, credits_total)."""
    journal_date, human = parse_date(date_str)
    description = f'Payroll {human}'
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    rows, warnings = [], []
    flagged = 0
    current_class, class_flagged = None, False
    total_dr = total_cr = Decimal('0')

    for row_no, cells in enumerate(ws.iter_rows(values_only=True), start=1):
        cells = list(cells) + [None] * (4 - len(cells))
        first = cells[0]
        text = str(first).strip() if first is not None else ''
        if not text and cells[1] is None and cells[2] is None:
            continue
        low = text.lower()
        if 'subtotal' in low:
            current_class, class_flagged = None, False
            continue
        if low.startswith(NOISE_PREFIXES) or NUMBERED_NOTE.match(text) or low == 'account':
            continue
        if low.startswith(('closing entries', 'grand total', 'department grand total', 'combined grand total')):
            current_class, class_flagged = None, False
            continue
        if 'total' in low and cells[1] is None and cells[2] is None:
            continue

        has_amount = cells[1] is not None or cells[2] is not None
        if not has_amount:
            if DEPT_HEADER.match(text):
                current_class, class_flagged = normalize_class(text, row_no, warnings)
            # any other amount-less text row is layout noise
            continue

        # ---- account line ----
        row_flag = False
        try:
            debit = to_amount(cells[1]); debit_s = fmt(debit)
        except (InvalidOperation, ValueError):
            debit, debit_s, row_flag = None, review(str(cells[1])), True
            warnings.append(f"WARNING: non-numeric debit {cells[1]!r} at row {row_no} — using placeholder, fix before import")
        try:
            credit = to_amount(cells[2]); credit_s = fmt(credit)
        except (InvalidOperation, ValueError):
            credit, credit_s, row_flag = None, review(str(cells[2])), True
            warnings.append(f"WARNING: non-numeric credit {cells[2]!r} at row {row_no} — using placeholder, fix before import")

        if not row_flag and (debit or Decimal('0')) == 0 and (credit or Decimal('0')) == 0:
            continue    # $0 line: skip entirely
        if debit is not None and credit is not None and debit != 0 and credit != 0:
            warnings.append(f"WARNING: both debit and credit populated at row {row_no} — flagged, fix before import")
            credit_s, row_flag = review(f'{credit_s} both sides populated'), True
            credit = None
        if debit is not None and debit == 0:
            debit_s = ''
        if credit is not None and credit == 0:
            credit_s = ''

        account = ACCOUNT_NAME_MAP.get(text)
        if account is None:
            warnings.append(f"WARNING: unmapped account '{text}' at row {row_no} — using placeholder, fix before import")
            account, row_flag = review(text), True

        cls = current_class or ''
        if class_flagged:
            row_flag = True

        if debit:
            total_dr += debit
        if credit:
            total_cr += credit
        if row_flag:
            flagged += 1
        rows.append([journal_no, journal_date, '', account, debit_s, credit_s, description, '', '', cls])

    wb.close()
    return rows, warnings, flagged, total_dr, total_cr


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('input_xlsx')
    p.add_argument('--journal-no', required=True)
    p.add_argument('--date', required=True, help='pay date, D/M/YYYY (e.g. 1/8/2026)')
    p.add_argument('--output', required=True)
    a = p.parse_args(argv)

    try:
        rows, warnings, flagged, dr, cr = convert(a.input_xlsx, a.journal_no, a.date)
    except ValueError as e:
        print(f'ERROR: {e}', file=sys.stderr)
        return 2

    for w in warnings:
        print(w, file=sys.stderr)

    with open(a.output, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(CSV_COLUMNS)
        w.writerows(rows)

    rc = 0
    diff = dr - cr
    if abs(diff) > CENT:
        print(f'WARNING: OUT OF BALANCE — Debits: {dr:.2f}  Credits: {cr:.2f}  Diff: {diff:.2f}', file=sys.stderr)
        rc = 1
    else:
        print(f'Balanced: Debits = Credits = ${dr:,.2f}')
    print(f'{len(rows)} rows written to {a.output}; {flagged} flagged for review')
    if flagged:
        print(f'{flagged} row(s) contain <<REVIEW:>> placeholders — fix before importing to QuickBooks', file=sys.stderr)
        rc = 1
    return rc


if __name__ == '__main__':
    sys.exit(main())
