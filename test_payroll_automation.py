#!/usr/bin/env python3
"""Tests for payroll_automation.py against test_data/expected_results.json.
Run: python -m unittest test_payroll_automation -v
"""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from decimal import Decimal

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, 'test_data')
sys.path.insert(0, HERE)
import payroll_automation as pa  # noqa: E402
import payroll_je  # noqa: E402

with open(os.path.join(DATA, 'expected_results.json')) as f:
    EXPECTED = json.load(f)


def D(x):
    return Decimal(str(x)).quantize(Decimal('0.01'))


class BalancedCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.exp = EXPECTED['test_case_1_balanced']
        cls.tmp = tempfile.mkdtemp()
        cls.res, cls.files = pa.run(os.path.join(DATA, 'test_case_1_balanced.csv'), cls.tmp, quiet=True)

    def test_header_and_no_failures(self):
        self.assertEqual(self.res['header_date'], self.exp['pay_period_header'])
        self.assertEqual(self.res['failures'], [])
        self.assertEqual(self.res['flags'], [])
        self.assertIsNotNone(self.files)

    def test_department_matrix_every_figure(self):
        self.assertEqual(list(self.res['matrix']), list(self.exp['department_matrix']))
        for code, cats in self.exp['department_matrix'].items():
            for acct, (json_key, _qb) in pa.CATEGORY_META.items():
                self.assertEqual(self.res['matrix'][code][acct], D(cats[json_key]), (code, json_key))
        self.assertEqual(self.res['matrix_total'], D(self.exp['department_matrix_grand_total']))

    def test_closing_entries(self):
        for key, val in self.exp['closing_entries'].items():
            self.assertEqual(self.res['closing'][key], D(val), key)
        self.assertEqual(self.res['closing_total'], D(self.exp['closing_entries_total']))
        self.assertEqual(self.res['closing_total'], self.res['matrix_total'])

    def test_legs(self):
        for key, val in self.exp['leg_a_dues_gst_pst'].items():
            self.assertEqual(self.res['leg_a'][key], D(val), key)
        for key, val in self.exp['leg_b_ei_cpp_fedtax'].items():
            self.assertEqual(self.res['leg_b'][key], D(val), key)

    def test_cross_checks(self):
        cc = self.exp['cross_checks']
        self.assertEqual(self.res['category_totals'][6046],
                         D(cc['cpp_expense_matrix_total_must_equal_top_block_cpp_employer_total']['matrix_side']))
        self.assertEqual(self.res['category_totals'][6046], self.res['cpp_employer_top'])
        self.assertEqual(self.res['category_totals'][6047],
                         D(cc['ei_expense_matrix_total_must_equal_top_block_ei_employer']['matrix_side']))
        self.assertEqual(self.res['ei_employer_top'],
                         D(cc['ei_expense_matrix_total_must_equal_top_block_ei_employer']['top_block_side']))

    def test_health_benefits_zero_is_noted_not_flagged(self):
        self.assertEqual(self.res['closing']['6043_employer_health_benefits'], Decimal('0'))
        self.assertTrue(any('confirmed $0' in n for n in self.res['notes']))

    def test_workbooks_written_with_formulas(self):
        from openpyxl import load_workbook
        ref, draft = self.files
        self.assertTrue(os.path.basename(ref) == 'PYROLL_REFERENCE_20260901.xlsx')
        self.assertTrue(os.path.basename(draft) == '20260901_by_department_DRAFT.xlsx')
        ws = load_workbook(ref).active
        self.assertEqual(str(ws['A1'].value), '20260901')
        formulas = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('=')]
        self.assertGreater(len(formulas), 30)
        ws2 = load_workbook(draft).active
        self.assertIn('DRAFT — built from raw payroll CSV, not a posted JE.', ws2['A2'].value)

    def test_end_to_end_draft_to_qb_csv(self):
        """DRAFT workbook -> payroll_je.py must produce a balanced, fully-mapped CSV."""
        _ref, draft = self.files
        rows, warnings, flagged, dr, cr = payroll_je.convert(draft, 'JJ0001', '5/9/2026')
        self.assertEqual(warnings, [])
        self.assertEqual(flagged, 0)
        self.assertEqual(dr, cr)
        self.assertEqual(dr, D(self.exp['department_matrix_grand_total']))
        classes = {r[9] for r in rows if r[4]}
        self.assertEqual(classes, {'0010-ACCOMODATIONS', '0015-HOUSEKEEPING'})
        self.assertTrue(all(r[9] == '' for r in rows if r[5]))

    def test_cli_exit_code(self):
        p = subprocess.run([sys.executable, os.path.join(HERE, 'payroll_automation.py'),
                            os.path.join(DATA, 'test_case_1_balanced.csv'), '--out-dir', self.tmp],
                           capture_output=True, text=True)
        self.assertEqual(p.returncode, 0, p.stderr)

    def test_bank_mismatch_blocks_output(self):
        tmp = tempfile.mkdtemp()
        res, files = pa.run(os.path.join(DATA, 'test_case_1_balanced.csv'), tmp,
                            bank={'leg_a': '525.50', 'leg_b': '12160.00', 'net_pay': '14700.00'}, quiet=True)
        self.assertIsNone(files)
        self.assertTrue(any('CHECK 6' in f for f in res['failures']))
        res, files = pa.run(os.path.join(DATA, 'test_case_1_balanced.csv'), tmp,
                            bank={'leg_a': '525.50', 'leg_b': '12160.00', 'net_pay': '14500.00'}, quiet=True)
        self.assertIsNotNone(files)
        self.assertTrue(res['bank_confirmed'])


class UnknownsCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.exp = EXPECTED['test_case_2_unknowns']['expected_flags']
        cls.tmp = tempfile.mkdtemp()
        cls.res, cls.files = pa.run(os.path.join(DATA, 'test_case_2_unknowns.csv'), cls.tmp, force=True, quiet=True)

    def test_flags_match_expected(self):
        got = {(f['type'], f['department_code'], f.get('label'), f.get('amount')) for f in self.res['flags']}
        want = {(e['type'], e['department_code'], e.get('label'),
                 D(e['amount']) if 'amount' in e else None) for e in self.exp}
        self.assertEqual(got, want)

    def test_placeholders_never_guess(self):
        self.assertEqual(self.res['dept_names']['99'], 'DEPT 99 (name not confirmed)')
        self.assertEqual(self.res['unmapped'], [('99', 'Mystery Allowance', Decimal('-500.00'))])
        # not folded into wages or anything else
        self.assertEqual(sum(self.res['matrix']['99'].values()), Decimal('0'))

    def test_validation_blocks_without_force(self):
        res, files = pa.run(os.path.join(DATA, 'test_case_2_unknowns.csv'), tempfile.mkdtemp(), quiet=True)
        self.assertIsNone(files)
        self.assertTrue(res['failures'])

    def test_draft_shows_placeholders_highlighted(self):
        from openpyxl import load_workbook
        ws = load_workbook(self.files[1]).active
        cells = [c for row in ws.iter_rows() for c in row if c.value is not None]
        hdr = next(c for c in cells if c.value == '99 — DEPT 99 (name not confirmed)')
        self.assertEqual(hdr.fill.fgColor.rgb[-6:], 'FFC7CE')
        unm = next(c for c in cells if c.value == pa.UNMAPPED_MARK)
        self.assertEqual(unm.fill.fgColor.rgb[-6:], 'FFC7CE')
        notes = [c.value for c in cells if isinstance(c.value, str) and 'Mystery Allowance' in c.value and 'NEEDS' not in c.value]
        self.assertTrue(any("department 99" in n and '-500.00' in n for n in notes))
        ws_ref = load_workbook(self.files[0]).active
        self.assertTrue(any(c.value == '99 — DEPT 99 (name not confirmed)' and c.fill.fgColor.rgb[-6:] == 'FFC7CE'
                            for row in ws_ref.iter_rows() for c in row))


if __name__ == '__main__':
    unittest.main()
