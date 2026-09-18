# Task: Build a Payroll Journal Entry CSV Generator

Build one Python script, `payroll_je.py`, that converts a "DRAFT by-department"
payroll workbook (.xlsx) into a QuickBooks Online journal-entry import CSV.
No web calls, no LLM calls inside the script — pure deterministic parsing.
Keep it to a single file, stdlib + openpyxl only. Optimize for low token/credit
use: write the whole file in one pass using the spec below rather than
iterating interactively; only re-open the file to fix a real bug.

## CLI

```
python payroll_je.py <input.xlsx> --journal-no JJ1234 --date DD/MM/YYYY --output <out.csv>
```

- `--journal-no` and `--date` are required arguments (no defaults — these
  change every pay period and must never be silently inferred).
- `--date` must be validated as D/M/YYYY. Reject and error out on MM/DD/YYYY-
  looking input (e.g. if day > 12, it's almost certainly fine; if the user
  passes an ambiguous date, still require explicit D/M/YYYY and do not guess).

## Input shape

The input workbook has ONE sheet. It is a human-edited draft with this layout
(row order may drift release to release — parse by content pattern, not fixed
row numbers):

- Department section headers, format: `"<code> — <CLASS NAME>"`, e.g.
  `"10 — 0010-ACCOMODATIONS"`. Sometimes annotated with a trailing
  parenthetical like `"(Salaried)"` or `"(name not confirmed)"` — strip
  anything from the first `" ("` onward when extracting the class name.
- Under each header, a small table with columns: Account | Debit | Credit | Description.
  A literal header row `"Account"` may appear — skip it.
- Department sections end at a line containing "Subtotal" (case-insensitive).
- After all departments, a "Closing entries" or "GRAND TOTAL" section holds
  company-wide credit lines (payables) and any company-wide debit lines
  (e.g. Utilities, Employer Health Benefits). These carry NO class.
- Noise to skip outright (line startswith, case-insensitive): "COMPANY-WIDE",
  "NOTES", "DRAFT", "Payroll Journal", "Check", numbered note lines ("1.",
  "2." ...), blank rows.

## Account name normalization — DO NOT GUESS

QuickBooks requires an EXACT string match against the existing Chart of
Accounts. Trailing dashes are inconsistent and NOT predictable from the
account number pattern — this cost real failed imports before. Use this
fixed lookup table as the single source of truth. Any account name found in
the workbook that is NOT a key in this table (after trimming whitespace)
must NOT be guessed or silently passed through — see "Unknown data" below.

```python
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
```

Matching should be exact after `.strip()` — do not fuzzy-match, do not
strip/add dashes heuristically.

## Class code normalization

The department header gives you the class name directly (e.g.
`0010-ACCOMODATIONS`) — use it verbatim once the parenthetical annotation is
stripped, EXCEPT:

```python
CLASS_NAME_FIXUPS = {
    "0097-INFORMATION TECHNOLOGY": "0097-INFORMATION TECHNOLOGY (IT)",
}
```

Apply this fixup map the same exact-match way as account names.

## Unknown data → placeholder + flag, never guess

If the script encounters:
- an account name not in `ACCOUNT_NAME_MAP`,
- a department header whose class, after fixups, doesn't look like the
  `NNNN-NAME` pattern, or is empty/unparseable,
- a debit/credit cell that isn't numeric,

it must NOT invent a value. Instead:
1. Write the row to the output CSV anyway, using `"<<REVIEW: original text>>"`
   in place of the bad field (so the row visibly fails QB's import and gets
   caught before posting, rather than being silently wrong or silently
   dropped).
2. Print a one-line warning to stderr for each flagged row:
   `WARNING: unmapped account 'X' at row N — using placeholder, fix before import`
3. At the end, print a summary count of flagged rows. If any exist, exit with
   a non-zero exit code (but still write the CSV — partial output is more
   useful than none for a human to patch).

## Output CSV

Columns, in order:
```
*JournalNo,*JournalDate,Memo,*AccountName,Debits,Credits,Description,Name,Location,Class
```

- `*JournalNo` and `*JournalDate` repeated on **every** row (QB rejects the
  file otherwise — confirmed failure mode).
- `Memo`, `Name`, `Location` are always blank.
- `Description` = `f"Payroll {date_human}"` where `date_human` is a plain
  English rendering of the pay date, e.g. `"Payroll 01 August 2026"` — derive
  it from the `--date` argument, don't ask the user to also type this.
- `Debits`/`Credits`: plain decimals, 2dp, no `$`, no thousands separator.
  Exactly one of the two populated per row, never both.
- Rows with debit AND credit both zero/blank are skipped entirely (a $0
  line, e.g. Employer Health Benefits some periods, should not appear in
  the output — confirmed correct behavior from a real prior period).
- `Class` blank for closing-entry/company-wide rows; populated for
  department rows.

## Balance check — mandatory, not optional

Before writing the file, sum all Debits and all Credits. If they differ by
more than $0.01:
- still write the CSV (so it can be inspected),
- print `WARNING: OUT OF BALANCE — Debits: X  Credits: Y  Diff: Z` to stderr,
- exit non-zero.
If they match, print `Balanced: Debits = Credits = $X` to stdout.

## Efficiency constraints

- Single pass over the worksheet rows (`openpyxl`, `data_only=True`,
  `iter_rows(values_only=True)`). Don't re-read the file multiple times.
- No pandas dependency needed for this — plain openpyxl keeps it lighter.
- No network calls, no subprocess calls, no interactive prompts mid-run —
  everything needed comes from CLI args + the workbook.

## Test fixture

Also generate `test_fixture.xlsx` (a small synthetic workbook, 3 departments
+ closing entries, values chosen so debits = credits exactly) and a
`test_payroll_je.py` using stdlib `unittest` that:
1. Runs the script against the fixture with `--journal-no TEST001
   --date 1/1/2026`.
2. Asserts the output CSV balances (sum debits == sum credits).
3. Asserts `*JournalNo` appears on every data row.
4. Asserts every `*AccountName` in the output is a value from
   `ACCOUNT_NAME_MAP`, never a raw un-mapped key.
5. Includes ONE deliberately-unmapped account name in the fixture (e.g.
   `"9999 Made Up Account"`) and asserts the script flags it with a
   `<<REVIEW:` placeholder and exits non-zero — this proves the "don't
   guess" behavior actually works, not just the happy path.

Keep the fixture data small (3 departments × ~4 line items + 4-5 closing
lines is enough) — don't recreate a full 18-department real payroll for the
test.

## Deliverables

- `payroll_je.py`
- `test_fixture.xlsx`
- `test_payroll_je.py`
- Run the test suite once at the end and paste the pass/fail output.
