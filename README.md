# Payroll Process

Two deterministic scripts (stdlib + openpyxl, no LLM or network calls) that take a
raw Payworks payroll CSV all the way to a QuickBooks Online journal-entry import.

```
raw payroll CSV
   │  python payroll_automation.py <raw.csv> --out-dir out/
   ├── PYROLL_REFERENCE_<period>.xlsx      reference workpaper (live formulas)
   └── <period>_by_department_DRAFT.xlsx   QB-ready department JE draft
          │  python payroll_je.py <DRAFT.xlsx> --journal-no JJ1234 --date D/M/YYYY --output je.csv
          └── je.csv                       QuickBooks Online JE import file
```

## Step 1: `payroll_automation.py`

```
python payroll_automation.py test_data/test_case_1_balanced.csv --out-dir out/
python payroll_automation.py raw.csv --out-dir out/ --bank-leg-a 525.50 --bank-leg-b 12160.00 --bank-net-pay 14500.00
```

* All GL mappings and formulas are constants at the top of the file (spec: `docs/PAYROLL_AUTOMATION_SPEC.md`).
* Validation checks 1 to 6 run before anything is written. On failure the diagnostic is
  printed, nothing is written, exit code 2. Pass `--force` to write anyway after review.
* Unknown department codes and unmapped pay-component labels are never guessed. They are
  highlighted red, listed under NOTES / NEEDS REVIEW, and the script exits 1.
* Bank figures are optional. Supplied and matching marks lines "confirmed against bank
  statement"; a mismatch is validation check 6 and blocks output.

## Step 2: `payroll_je.py`

```
python payroll_je.py out/20260901_by_department_DRAFT.xlsx --journal-no JJ2700 --date 5/9/2026 --output je.csv
```

* `--journal-no` and `--date` are required. `--date` is strict D/M/YYYY.
* Account names are mapped through `ACCOUNT_NAME_MAP` by exact match. Anything unmapped is
  written as `<<REVIEW: ...>>`, warned on stderr, and the script exits 1 (the CSV is still written).
* Debits and credits are summed before writing. Out of balance also exits 1.

## Tests

```
pip install -r requirements.txt
python make_test_fixture.py          # regenerates test_fixture.xlsx if needed
python -m unittest test_payroll_je test_payroll_automation -v
```

`test_data/` holds the hand-verified fixtures and `expected_results.json`; every figure in
that file is asserted exactly, plus an end-to-end test that feeds the generated DRAFT
workbook into `payroll_je.py` and checks the resulting CSV balances with zero flags.
