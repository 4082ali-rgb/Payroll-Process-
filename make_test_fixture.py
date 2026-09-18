#!/usr/bin/env python3
"""Generate test_fixture.xlsx for test_payroll_je.py.

3 departments + closing entries; debits == credits exactly (7,000.00 each side).
Includes ONE deliberately unmapped account ("9999 Made Up Account") and one $0
line that must be skipped. Layout mimics a hand-edited DRAFT workbook, including
noise rows and a (Salaried) annotated header.
"""
from openpyxl import Workbook

ROWS = [
    ["Payroll Journal Entry — batch 20260101 — by department"],
    ["DRAFT — built from raw payroll CSV, not a posted JE."],
    [],
    ["10 — 0010-ACCOMODATIONS"],
    ["Account", "Debit", "Credit", "Description"],
    ["6050 Wages & Salaries", 2000.00, None, "Accommodations wages"],
    ["6046 CPP Expense", 100.00, None, ""],
    ["6047 EI Expense", 50.00, None, ""],
    ["6049 Vacation expense", 80.00, None, ""],
    ["Subtotal", "=SUM(B6:B9)"],
    [],
    ["0020-1 — 0020-PINEWOODS (Salaried)"],
    ["Account", "Debit", "Credit", "Description"],
    ["6050 Wages & Salaries", 1500.00, None, ""],
    ["6046 CPP Expense", 90.00, None, ""],
    ["6047 EI Expense", 40.00, None, ""],
    ["9999 Made Up Account", 25.00, None, "deliberately unmapped"],
    ["Subtotal", "=SUM(B14:B17)"],
    [],
    ["97 — 0097-INFORMATION TECHNOLOGY"],
    ["Account", "Debit", "Credit", "Description"],
    ["6050 Wages & Salaries", 2800.00, None, ""],
    ["6046 CPP Expense", 200.00, None, ""],
    ["6047 EI Expense", 100.00, None, ""],
    ["6044 Tips & Gratuities", 15.00, None, ""],
    ["Subtotal", "=SUM(B22:B25)"],
    [],
    ["Department grand total", "=B10+B18+B26"],
    [],
    ["Company-wide closing entries"],
    ["Account", "Debit", "Credit", "Description"],
    ["2016 Wages Payable", None, 4500.00, "not yet confirmed"],
    ["2017 Federal Income Tax Payable", None, 1200.00, ""],
    ["2014 CPP Payable", None, 800.00, ""],
    ["2013 EI Payable", None, 400.00, ""],
    ["2015 RRSP Payable", None, 100.00, ""],
    ["6043 Employer Health Benefits", None, 0.00, "known-$0 period"],
    ["Closing total", None, "=SUM(C32:C37)"],
    [],
    ["Combined grand total", "=B28", "=C38"],
    ["Check (Debit - Credit, must be $0.00)", "=B28-C38"],
    [],
    ["NOTES — READ BEFORE POSTING"],
    ["1. Employer Health Benefits confirmed $0 this period."],
    ["2. Bank confirmation: not yet confirmed."],
]


def build(path="test_fixture.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "DRAFT JE"
    for r in ROWS:
        ws.append(r)
    wb.save(path)
    return path


if __name__ == "__main__":
    print("wrote", build())
