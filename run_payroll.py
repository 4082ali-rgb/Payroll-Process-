#!/usr/bin/env python3
"""One command: raw Payworks payroll CSV -> the four payroll deliverables.

Usage:
    python run_payroll.py <raw_payroll.csv> --bank-date D/M/YYYY \\
        --dept-journal-no JJ1 --dues-journal-no JJ2 --ei-journal-no JJ3 \\
        [--out-dir DIR] [--bank-leg-a 1252.16] [--bank-leg-b 51475.50] [--bank-net-pay 137803.11]

Outputs (in --out-dir, default: current directory):
    PYROLL_REFERENCE_<period>.xlsx     1. worksheet (live formulas, checks)
    <period>_by_department_DRAFT.xlsx  intermediate draft the department journal is built from
    <JJ>_Department_Breakdown.csv      2. journal, Memo "Department Breakdown"
    <JJ>_Dues_and_Subscriptions.csv    3. Leg A journal, Memo "Dues and Subscriptions"
    <JJ>_EI_CPP_Federal_Tax.csv        4. Leg B journal, Memo "EI, CPP & Federal Tax"

--bank-date is the day the entries leave the bank (not the batch header date) and is
used as the journal date on all three journals. The three journal numbers and the bank
date are required: they change every pay period and are never guessed.

Deterministic. No LLM, no network. Exit code is the worst of the steps (0 = clean).
"""

import argparse
import os
import sys

import leg_journal_csv
import payroll_automation
import payroll_je


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('csv_path')
    p.add_argument('--bank-date', required=True, help='date the entries leave the bank, D/M/YYYY')
    p.add_argument('--dept-journal-no', required=True)
    p.add_argument('--dues-journal-no', required=True)
    p.add_argument('--ei-journal-no', required=True)
    p.add_argument('--out-dir', default='.')
    p.add_argument('--bank-leg-a', help='bank withdrawal for Dues/GST/PST')
    p.add_argument('--bank-leg-b', help='bank withdrawal for EI/CPP/Federal Tax')
    p.add_argument('--bank-net-pay', help='bank withdrawal for Net Pay')
    p.add_argument('--force', action='store_true', help='write files even if validation fails')
    a = p.parse_args(argv)

    try:
        payroll_je.parse_date(a.bank_date)  # fail fast, before anything is written
    except ValueError as e:
        print(f'ERROR: {e}', file=sys.stderr)
        return 2
    os.makedirs(a.out_dir, exist_ok=True)

    bank = None
    if a.bank_leg_a or a.bank_leg_b or a.bank_net_pay:
        bank = {'leg_a': a.bank_leg_a, 'leg_b': a.bank_leg_b, 'net_pay': a.bank_net_pay}
    res, files = payroll_automation.run(a.csv_path, a.out_dir, bank=bank, force=a.force)
    if files is None:
        print('Validation failed: no files written.', file=sys.stderr)
        return 2
    rc = 2 if res['failures'] else (1 if res['flags'] else 0)
    if res['flags']:
        print(f"{len(res['flags'])} item(s) need review - see NOTES section in the DRAFT workbook.", file=sys.stderr)
    draft = files[1]

    dept_out = os.path.join(a.out_dir, f'{a.dept_journal_no}_Department_Breakdown.csv')
    dues_out = os.path.join(a.out_dir, f'{a.dues_journal_no}_Dues_and_Subscriptions.csv')
    ei_out = os.path.join(a.out_dir, f'{a.ei_journal_no}_EI_CPP_Federal_Tax.csv')

    rc = max(rc, payroll_je.main([draft, '--journal-no', a.dept_journal_no, '--date', a.bank_date,
                                  '--output', dept_out]))
    rc = max(rc, leg_journal_csv.main([a.csv_path, '--leg', 'a', '--journal-no', a.dues_journal_no,
                                       '--date', a.bank_date, '--output', dues_out]))
    rc = max(rc, leg_journal_csv.main([a.csv_path, '--leg', 'b', '--journal-no', a.ei_journal_no,
                                       '--date', a.bank_date, '--output', ei_out]))
    return rc


if __name__ == '__main__':
    sys.exit(main())
