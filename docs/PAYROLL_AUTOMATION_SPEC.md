# Payroll Journal Entry Automation — Build Spec

## Goal and efficiency principle

Automate turning a raw Payworks payroll CSV export into two output files:
1. A full reference workpaper (mirrors the existing `PYROLL_REFERENCE_*.xlsx` format).
2. A QuickBooks-ready department journal entry (Debit/Credit columns, copy-paste ready).

**Write this as one deterministic Python script, not an LLM-driven process.** Every GL
account, every formula, and every mapping below is already fixed business logic —
verified against two real posted QuickBooks journal entries (Jan 21 2026 "JJ2649" and
Aug 1 2026). None of it requires interpretation at runtime. The only place judgment is
ever needed is when the script meets something genuinely new (an unmapped department
code or pay-component label) — and even then the correct behavior is just "flag it,"
not "figure it out."

Once written, running this script for a new pay period should cost **zero additional
LLM calls** — it's pure `csv` + `openpyxl` execution. Do not re-derive the GL logic,
re-verify formulas, or re-read source files each run. Encode everything below as
constants/lookup tables once, at the top of the script.

Test fixtures with hand-verified expected output are in `test_data/`. Write real
`assert`-based tests against `expected_results.json` — do not eyeball the output.

---

## 1. Input format

Raw CSV, no header row for columns. Structure:

```
<row 1>                    YYYYMMDD (pay-period batch date, NOT the pay date — actual
                            bank withdrawal happens ~4 days later)
<row 2 onward>              Batch,DeptCode,,Label,Amount
```

- `Batch` is always `B10657` (or similar — treat as a passthrough field, don't hardcode).
- `DeptCode` blank = company-wide ("top block") line. Non-blank = department line.
- Department codes come as zero-padded numeric strings (`000010`, `000015`, `0010-1` for
  salaried). Strip leading zeros for numeric codes; keep the `-1` suffix on salaried
  codes as-is (don't strip those).
- The file contains company-wide lines **twice**: once at the start (the full top block,
  ending in a `Payroll Clearing Account` line), then all the department lines, then a
  **second, short company-wide block** at the end that just reverses Service Fees / GST
  / PST and repeats `Payroll Clearing Account` with the opposite sign. This closing
  block is bookkeeping noise from the source export — **ignore it**, don't double-count.
  Detect it by watching for `DeptCode` going blank again *after* department rows have
  already started.

### Parsing algorithm
```
mode = 'top'
for each data row:
    if DeptCode is blank:
        if mode == 'top': append to top_block
        else: mode = 'closing'; ignore (discard row)
    else:
        mode = 'dept'
        normalize dept code (strip leading zeros; keep '-1' suffix)
        append to dept_blocks[dept_code]  (preserve first-seen order in dept_order)
```

---

## 2. Department code → name mapping

```python
DEPT_NAMES = {
    '10': '0010-ACCOMODATIONS',
    '15': '0015-HOUSEKEEPING',
    '20': '0020-PINEWOODS',
    '25': '0025-DAYLODGE',            # seasonal — appears/disappears period to period
    '30': '0030-COUNTRY STORE',
    '40': '0040-BOAT HOUSE',
    '50': '0050-MANNING PARKS',
    '51': '0051-VISITOR CENTER',
    '52': '0052-SKYVIEW',
    '60': '0060-ALPINE',
    '70': '0070-NORDIC',              # seasonal
    '80': '0080-MAINTENANCE',
    '81': '0081-WEDDINGS & BANQUETS',
    '85': '0085-STAFF HOUSING',
    '90': '0090-GENERAL & ADMIN',
    '91': '0091-KITCHEN',
    '92': '0092-BEARS DEN',
    '93': '0093-MARKETING',
    '94': '0094-HUMAN RESOURCES',
    '95': '0095-ALPINE RENTALS',
    '96': '0096-ALPINE LESSONS',      # seasonal
    '97': '0097-INFORMATION TECHNOLOGY (IT)',
    '999': '0999-SKI PATROL',         # seasonal
    '0010-1': '0010-ACCOMODATIONS (Salaried)',
    '0020-1': '0020-PINEWOODS (Salaried)',
    '0050-1': '0050-MANNING PARKS (Salaried)',
    '0060-1': '0060-ALPINE (Salaried)',
}
```

**Any department code not in this table is unknown.** Departments 0, 1, 40, 51 have all
shown up at one time or another before being identified — 40 and 51 are now confirmed
above; 0 and 1 were one-off oddities (single tiny line items, no wages) never resolved.

**Required behavior for an unknown code:** label it `f"DEPT {code} (name not confirmed)"`,
highlight it (fill color, e.g. `FFC7CE` red-ish) in both output files, and list it in the
NEEDS REVIEW notes section. **Never guess a name.** This is non-negotiable — a wrong
guessed name is worse than an honest placeholder.

---

## 3. GL account rules — Department matrix (8 categories, ALL DEBIT)

Verified against real posted JEs: every department-level line, across all 8 categories
below, is a **debit**. None are credits. (This was gotten wrong once during manual work
and corrected after checking the real data — don't repeat that mistake.)

```python
GL_CATEGORY_SOURCE_LABELS = {
    (6050, 'WAGES & SALARIES'): [
        'Regular Pay', 'Regular Salary', 'Wages', 'Overtime', 'Sick Pay', 'Training',
        'Double Time', 'Retro Regular Pay', 'Stat Pay @1.0', 'Stat Pay @1.5',
    ],
    (6046, 'CPP EXPENSE'): ['CPP/QPP Employer', 'CPP2/QPP2 Employer'],
    (6047, 'EI EXPENSE'): ['EI Employer'],
    (6049, 'VACATION EXPENSE'): ['Vac Each Pay'],
    (6044, 'TIPS & GRATUITIES'): ['Tips & Gratitude'],
    (2024, 'STAFF DEPOSIT'): ['Refund Staff Deposit'],
    (6043, 'BENEFITS'): ['Benefits'],
    (2011, 'VACATION PAYABLE'): [],   # always 0 — placeholder/control row, no source data maps here
}
```

For each department, for each category, sum whichever of that category's source labels
actually appear for that department (most departments won't have all of them — that's
normal, use 0/omit, don't treat as an error).

**Any pay-component label encountered for a department that does not appear in ANY of
the source-label lists above is unmapped.** Do not silently fold it into Wages & Salaries
or drop it. Flag it: insert a `⚠️ UNMAPPED` placeholder in the GL account column, and
list the raw label, department, and amount in the NEEDS REVIEW section.

---

## 4. GL account rules — Top-block legs (two separate, self-contained JEs)

These are **not part of** the department matrix / closing-entry "big entry." They're
two smaller, independent journal entries that also draw from Envision Bank. Keep them
in their own section of the output, clearly separated.

**Leg A — Dues/GST/PST (all amounts from the top block only):**
```
5004 Dues and Subscription   = Service Fees                          [DEBIT]
2031 GST Paid on Purchases   = GST                                   [DEBIT]
2040 PST Paid on Purchases   = PST                                   [DEBIT]
1004 Envision Bank Account   = SUM(above three)                      [CREDIT]
```

**Leg B — EI/CPP/Federal Tax remittance:**
```
2013 EI Payable                    = EI Employee + EI Employer            [DEBIT]
2014 CPP Payable                   = CPP/QPP Employee + CPP2/QPP2 Employee (if present)
                                    + CPP/QPP Employer + CPP2/QPP2 Employer (if present)  [DEBIT]
2017 Federal Income Tax Payable    = Federal Tax                          [DEBIT]
1004 Envision Bank Account         = SUM(above three)                     [CREDIT]
```

Note: CPP2/QPP2 lines are present in some pay periods and absent in others (enhanced
CPP contribution, applies once earnings cross a threshold). Treat their absence as
normal, not an error — just don't include a term for a line that doesn't exist that
period.

---

## 5. GL account rules — Closing entries (8 categories, ALL CREDIT)

This is the "big entry" that offsets the department matrix. The department matrix total
(section 3, summed across every department and category) must equal this closing total
exactly — see section 6.

```
6007 Utilities                   = Utilities 100 + Utilities 170 (+ Utilities Repayment,
                                    if that line exists that period)                [CREDIT]
2024 Staff Deposit               = Staff Deposit (the COMPANY-WIDE top-block line —
                                    this is a DIFFERENT figure from the department-level
                                    "Refund Staff Deposit" lines in section 3; don't
                                    confuse them)                                    [CREDIT]
2016 Wages Payable               = Net Pay
                                    + Prov. Garnishment      (if present that period)
                                    + Net Pay Deduction      (if present that period)
                                    + Excess Deductions      (if present that period)  [CREDIT]
2017 Federal Income Tax Payable  = Federal Tax                                       [CREDIT]
2014 CPP Payable                 = same formula as Leg B                             [CREDIT]
2013 EI Payable                  = same formula as Leg B                             [CREDIT]
2015 RRSP Payable                = RRSP EE + RRSP ER                                 [CREDIT]
6043 Employer Health Benefits    = AD&D + Critical Illness + Dental + EAP
                                    + Extended Health + Life Insurance + LTD  (employee side)
                                    + AD&D ER + Critical Illness ER + Dental ER + EAP ER
                                    + Extended Health ER + Life Insurance ER + LTD ER
                                    (employer side)                                   [CREDIT]
```

**Important:** the seven AD&D/Dental/etc. deduction lines that feed Employer Health
Benefits are present in some pay periods and completely absent in others (they don't
appear to be deducted every single pay cycle). If none of them exist in the raw data
for a period, **Employer Health Benefits is genuinely $0.00** — this is correct, not a
gap. Don't flag it as missing; do flag it in the output notes as "confirmed $0 for this
period" so a human reviewing later knows it was checked, not skipped.

The **Wages Payable** formula is the one part of this spec built from four *possible*
components (Net Pay always present; Garnishment, Net Pay Deduction, and Excess
Deductions each independently present or absent depending on the period). Include only
the ones that exist in that period's raw data — check by label presence, not by
assuming a fixed set.

---

## 6. Validation checks — must run automatically, must block output on failure

Run these after computing everything. If any fails, **do not silently deliver the
output** — halt, print a clear diagnostic showing both sides of the failing check, and
require human review before the files are written.

1. **Department matrix grand total (magnitude) == closing entries total (magnitude)**,
   to the penny. This is the core double-entry check for the "big entry."
2. **CPP Expense matrix total == top-block CPP/QPP Employer + CPP2/QPP2 Employer (if
   present)**. Independent cross-check, unrelated to check #1's formula path.
3. **EI Expense matrix total == top-block EI Employer**. Same idea.
4. **Leg A debits == Leg A credit** (525.50-style three-line self-balance).
5. **Leg B debits == Leg B credit**.
6. If bank statement figures are supplied (see section 7), cross-check the three real
   withdrawal amounts against Leg A total, Leg B total, and the Net-Pay-only component
   of Wages Payable (not the full Wages Payable figure, which also includes garnishment/
   deduction amounts that don't move through the bank the same day).

If a validation check fails by a small, exact amount, look for a data-entry-style
explanation (a transposed digit, a value copied into the wrong row) before concluding
the formulas are wrong — in past manual runs, every discrepancy found this way turned
out to be exactly that, not a logic error.

---

## 7. Optional bank statement cross-check

If the person supplies bank statement figures for the period (three Envision
withdrawals dated ~4 days after the batch header date), match them against:
- Leg A total (Dues/GST/PST)
- Leg B total (EI/CPP/FedTax)
- The Net Pay component only (not the full Wages Payable figure — see section 5)

If supplied and matching, mark the relevant lines "confirmed against bank statement" in
the output notes. If not supplied, leave the bank-rec box blank/highlighted and note
"not yet confirmed" — don't fabricate a bank figure.

---

## 8. Output 1 — Reference workpaper (`PYROLL_REFERENCE_<period>.xlsx`)

Structure, top to bottom, one sheet:
1. Header cell: batch date (`YYYYMMDD` as entered, e.g. `20260901`).
2. Top block: every top-block line, label + amount, in original order.
3. Top-block GL coding (Leg A + Leg B from section 4), using formulas (`=E{row}`
   references), not hardcoded values — so the sheet stays self-auditing if source
   numbers change.
4. One block per department, in `dept_order` (first-seen order from the raw file):
   department header row, line items, then a `=SUM(...)` subtotal row.
5. Cross-tab matrix: one column per department, one row per GL category (section 3,
   8 rows), each cell a `=SUM(...)` formula referencing that department's matching
   source rows, plus a grand-total column.
6. Closing checklist (section 5), each line a formula referencing the top-block rows.
7. Bank rec box: `NetPay` (`=` top-block Net Pay cell), `Bank` (manual entry, highlighted
   yellow, blank if not yet known), `Difference` (`=NetPay-Bank`).

Use live formulas throughout, not computed constants — this is what makes the sheet
re-auditable by a human in Excel without re-running the script.

## 9. Output 2 — QB-ready department JE (`<period>_by_department_DRAFT.xlsx`)

Structure:
1. Title + a bold warning banner: "DRAFT — built from raw payroll CSV, not a posted JE."
   Update this banner if/when bank-confirmed (see section 7).
2. One block per department: header, Account/Debit/Credit/Description columns, one row
   per non-zero GL category from section 3 (all in the Debit column), `=SUM()` subtotal.
3. Department grand-total row (sum of all subtotals).
4. "Company-wide closing entries" section (section 5), all in the Credit column, each
   row's Description noting whether it's bank-confirmed, an internal-consistency-only
   figure, or a known-$0 period.
5. Combined grand total (Debit and Credit) and a `=Debit-Credit` check row that must
   read exactly `$0.00`.
6. A "NOTES — READ BEFORE POSTING" section listing, in order: any unknown department
   codes, any unmapped pay-component labels, any period-specific quirks (new CPP2 lines,
   missing benefit deductions, reappearing seasonal departments), and bank-confirmation
   status.

---

## 10. Test data

`test_data/test_case_1_balanced.csv` — a small, fully self-consistent synthetic payroll
batch (3 departments), hand-verified. `test_data/expected_results.json` has every
number the script must reproduce exactly, including the department matrix, closing
entries, both legs, and the two cross-checks. Write real `assert`s against this file —
if the script's output doesn't match every figure exactly, the build isn't done.

`test_data/test_case_2_unknowns.csv` — deliberately unrealistic and unbalanced. Its only
job is to confirm the script correctly flags an unknown department code (`99`) and an
unmapped pay-component label (`Mystery Allowance`) instead of guessing or silently
dropping them. Assert against the `expected_flags` list in the same JSON file — don't
assert a balance for this case, it's not meant to balance.

---

## 11. What NOT to do

- Don't call an LLM per pay period once this script exists — it's pure deterministic
  computation from here on.
- Don't guess a department name for an unknown code, ever.
- Don't silently fold an unmapped pay-component label into an existing GL category.
- Don't hardcode computed values into the output workbook — use live formulas.
- Don't assume Employer Health Benefits being $0 is an error — confirm via label
  presence/absence, then move on.
- Don't deliver output when a validation check (section 6) fails — surface it instead.
