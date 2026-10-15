#!/usr/bin/env python3
"""Generate a QuickBooks Online journal-entry import CSV for Leg A (Dues/GST/PST)
or Leg B (EI/CPP/Federal Tax remittance) from a raw Payworks payroll CSV.

These are the two small, self-contained entries that draw from the Envision Bank
Account alongside the main by-department journal entry (see docs/PAYROLL_AUTOMATION_SPEC.md
section 4). Each is 3 debit lines + 1 bank credit line that must self-balance.

Usage:
    python leg_journal_csv.py <raw_payroll.csv> --leg a --journal-no JJ1234 --date D/M/YYYY --output legA.csv
    python leg_journal_csv.py <raw_payroll.csv> --leg b --journal-no JJ1235 --date D/M/YYYY --output legB.csv

--journal-no and --date are required (no defaults - these change every pay period).
--date is strict D/M/YYYY, same rule as payroll_je.py.
--class defaults to "0090-GENERAL & ADMIN" (matches this company's actual posted
entries); pass --class "" for no class, per the original spec's "company-wide,
no class" assumption.

Deterministic. No LLM, no network. stdlib + openpyxl only.
"""

import argparse
import csv
import sys

from payroll_automation import parse_csv, compute, LEG_A, LEG_B, BANK_ACCOUNT
from payroll_je import parse_date, CSV_COLUMNS

DEFAULT_CLASS = '0090-GENERAL & ADMIN'


def build_rows(res, leg, journal_no, journal_date, description, class_name):
    leg_def, leg_values = (LEG_A, res['leg_a']) if leg == 'a' else (LEG_B, res['leg_b'])
    rows = []
    for key, name, _labels in leg_def:
        amt = leg_values[key]
        rows.append([journal_no, journal_date, '', name, f'{amt:.2f}', '', description, '', '', class_name])
    bank_amt = leg_values[BANK_ACCOUNT[0]]
    rows.append([journal_no, journal_date, '', BANK_ACCOUNT[1], '', f'{bank_amt:.2f}', description, '', '', ''])
    return rows


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('csv_path')
    p.add_argument('--leg', required=True, choices=['a', 'b'])
    p.add_argument('--journal-no', required=True)
    p.add_argument('--date', required=True, help='pay date, D/M/YYYY (e.g. 2/9/2026)')
    p.add_argument('--output', required=True)
    p.add_argument('--class', dest='class_name', default=DEFAULT_CLASS)
    a = p.parse_args(argv)

    try:
        journal_date, human = parse_date(a.date)
    except ValueError as e:
        print(f'ERROR: {e}', file=sys.stderr)
        return 2

    header_date, top_block, dept_blocks, dept_order, ignored = parse_csv(a.csv_path)
    res = compute(header_date, top_block, dept_blocks, dept_order, ignored=ignored)
    description = f'Payroll {human}'
    rows = build_rows(res, a.leg, a.journal_no, journal_date, description, a.class_name)

    debits = sum(float(r[4]) for r in rows if r[4])
    credits = sum(float(r[5]) for r in rows if r[5])

    with open(a.output, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(CSV_COLUMNS)
        w.writerows(rows)

    if abs(debits - credits) > 0.01:
        print(f'WARNING: OUT OF BALANCE — Debits: {debits:.2f}  Credits: {credits:.2f}  '
             f'Diff: {debits - credits:.2f}', file=sys.stderr)
        print(f'{len(rows)} rows written to {a.output}')
        return 1
    print(f'Balanced: Debits = Credits = ${debits:,.2f}')
    print(f'{len(rows)} rows written to {a.output}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
