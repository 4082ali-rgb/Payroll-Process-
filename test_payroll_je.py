#!/usr/bin/env python3
"""Tests for payroll_je.py (stdlib unittest). Run: python -m unittest test_payroll_je -v"""
import csv
import io
import os
import subprocess
import sys
import tempfile
import unittest
from decimal import Decimal

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, 'payroll_je.py')
FIXTURE = os.path.join(HERE, 'test_fixture.xlsx')

sys.path.insert(0, HERE)
import payroll_je  # noqa: E402
import make_test_fixture  # noqa: E402


def run_script(*args):
    return subprocess.run([sys.executable, SCRIPT, *args], capture_output=True, text=True)


class FixtureRun(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not os.path.exists(FIXTURE):
            make_test_fixture.build(FIXTURE)
        cls.tmp = tempfile.mkdtemp()
        cls.out = os.path.join(cls.tmp, 'out.csv')
        cls.proc = run_script(FIXTURE, '--journal-no', 'TEST001', '--date', '1/1/2026', '--output', cls.out)
        with open(cls.out, newline='', encoding='utf-8') as f:
            cls.rows = list(csv.DictReader(f))

    def test_header_columns(self):
        with open(self.out, newline='', encoding='utf-8') as f:
            header = next(csv.reader(f))
        self.assertEqual(header, payroll_je.CSV_COLUMNS)

    def test_balances(self):
        dr = sum(Decimal(r['Debits']) for r in self.rows if r['Debits'] and not r['Debits'].startswith('<<'))
        cr = sum(Decimal(r['Credits']) for r in self.rows if r['Credits'] and not r['Credits'].startswith('<<'))
        self.assertEqual(dr, cr)
        self.assertEqual(dr, Decimal('7000.00'))

    def test_journal_no_and_date_on_every_row(self):
        self.assertTrue(self.rows)
        for r in self.rows:
            self.assertEqual(r['*JournalNo'], 'TEST001')
            self.assertEqual(r['*JournalDate'], '01/01/2026')
            self.assertEqual(r['Description'], 'Payroll 01 January 2026')
            self.assertEqual(r['Memo'], ''); self.assertEqual(r['Name'], ''); self.assertEqual(r['Location'], '')

    def test_account_names_are_mapped_values_or_review(self):
        allowed = set(payroll_je.ACCOUNT_NAME_MAP.values())
        raw_keys_only = set(payroll_je.ACCOUNT_NAME_MAP) - allowed
        for r in self.rows:
            name = r['*AccountName']
            self.assertTrue(name in allowed or name.startswith('<<REVIEW:'), name)
            self.assertNotIn(name, raw_keys_only)

    def test_unmapped_account_flagged_and_nonzero_exit(self):
        flagged = [r for r in self.rows if r['*AccountName'].startswith('<<REVIEW:')]
        self.assertEqual(len(flagged), 1)
        self.assertEqual(flagged[0]['*AccountName'], '<<REVIEW: 9999 Made Up Account>>')
        self.assertIn("WARNING: unmapped account '9999 Made Up Account'", self.proc.stderr)
        self.assertNotEqual(self.proc.returncode, 0)
        self.assertIn('Balanced', self.proc.stdout)   # balance still reported

    def test_exactly_one_side_per_row_and_zero_rows_skipped(self):
        for r in self.rows:
            self.assertNotEqual(bool(r['Debits']), bool(r['Credits']), r)
        names = [r['*AccountName'] for r in self.rows]
        self.assertNotIn('6043 Employee Health Benefits', names)   # $0 line skipped

    def test_classes(self):
        by_name = {}
        for r in self.rows:
            by_name.setdefault(r['Class'], []).append(r)
        dept_rows = [r for r in self.rows if r['Debits']]
        closing_rows = [r for r in self.rows if r['Credits']]
        self.assertTrue(all(r['Class'] for r in dept_rows))
        self.assertTrue(all(r['Class'] == '' for r in closing_rows))
        self.assertIn('0010-ACCOMODATIONS', by_name)
        self.assertIn('0020-PINEWOODS', by_name)           # "(Salaried)" stripped
        self.assertIn('0097-INFORMATION TECHNOLOGY (IT)', by_name)  # fixup applied
        self.assertEqual(len(dept_rows), 12)
        self.assertEqual(len(closing_rows), 5)


class DateValidation(unittest.TestCase):
    def test_accepts_dmy(self):
        self.assertEqual(payroll_je.parse_date('1/8/2026'), ('01/08/2026', '01 August 2026'))
        self.assertEqual(payroll_je.parse_date('21/01/2026')[1], '21 January 2026')

    def test_rejects_mdy_looking(self):
        with self.assertRaises(ValueError):
            payroll_je.parse_date('08/21/2026')

    def test_rejects_other_formats(self):
        for bad in ('2026-08-01', '1 Aug 2026', '', '1/8/26'):
            with self.assertRaises(ValueError):
                payroll_je.parse_date(bad)

    def test_cli_requires_args(self):
        p = run_script(FIXTURE, '--output', 'x.csv')
        self.assertNotEqual(p.returncode, 0)
        p = run_script(FIXTURE, '--journal-no', 'X', '--date', '13/13/2026', '--output', os.devnull)
        self.assertEqual(p.returncode, 2)
        self.assertIn('MM/DD/YYYY', p.stderr)


class UnknownClassAndAmounts(unittest.TestCase):
    def test_bad_class_and_non_numeric_amount(self):
        from openpyxl import Workbook
        wb = Workbook(); ws = wb.active
        ws.append(['99 — DEPT 99 (name not confirmed)'])
        ws.append(['6050 Wages & Salaries', 100, None])
        ws.append(['Subtotal', 100])
        ws.append(['10 — 0010-ACCOMODATIONS'])
        ws.append(['6046 CPP Expense', 'abc', None])
        ws.append(['Subtotal'])
        ws.append(['Closing entries'])
        ws.append(['2016 Wages Payable', None, 100])
        tmp = tempfile.mkdtemp(); path = os.path.join(tmp, 'bad.xlsx'); wb.save(path)
        rows, warnings, flagged, dr, cr = payroll_je.convert(path, 'J1', '2/2/2026')
        self.assertEqual(rows[0][9], '<<REVIEW: 99 — DEPT 99 (name not confirmed)>>')
        self.assertEqual(rows[1][4], '<<REVIEW: abc>>')
        self.assertEqual(rows[1][9], '0010-ACCOMODATIONS')
        self.assertEqual(rows[2][9], '')
        self.assertEqual(flagged, 2)
        self.assertTrue(any('unparseable class' in w for w in warnings))
        self.assertTrue(any('non-numeric debit' in w for w in warnings))


if __name__ == '__main__':
    unittest.main()
