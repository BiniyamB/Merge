# POS & ATM & QR Report Merger

A small web dashboard that merges **any number of Excel reports** into
**one clean Excel report**, in five modes:

- **POS Decline** → merged in the `POS_Transaction_Decline_Report` layout
- **POS Success** → merged in the `POS_Transaction_SVFE_Report` layout
- **POS** → merged in the `Daily_Tranaction_Report_SmartVista_POS` layout
- **ATM** → merged in the `Daily_Tranaction_Report_SmartVista_ATM` layout
- **QR** → transfer exports merged in the `QR_Export` layout (the
  `EXPORT_TABLE` structure of the "July - December 2025 Source" report)

When you open the app it first asks which type of report you want to merge
(a **POS Decline**, **POS Success**, **POS**, **ATM** or **QR** button); the whole workflow — upload,
blank-column removal, header/reshuffle checks, missing/extra-column
warnings, preview and download — then runs in that mode. A breadcrumb bar
(`Home › POS Decline` …) lets you jump back to the mode picker at any time.
Blank/spacer columns are always removed, and uploaded files and the merged
result are processed **entirely in memory** — nothing is ever written to
disk.

## The output-sample rule

In every mode, the **sample output report defines the exact columns of the
merged result**:

- every uploaded file must have **at least** the sample's columns — if one
  is missing a column, a warning names the missing column **and** the file
  it is missing from
- any **additional** column in an uploaded file that is *not* in the sample
  (e.g. a `UTRNNO` column in a POS export whose sample has none) is simply
  **removed** — the merged output contains only the sample's columns, and
  the removed columns are noted per file on the dashboard

## Features

- On opening, a mode picker asks for **POS Decline**, **POS Success**, **POS**, **ATM** or **QR**
  reports; a breadcrumb bar (Home > mode > results) lets you go back to the
  picker
- Drag & drop any number of `.xls` / `.xlsx` reports (or browse for them)
- Automatically detects the report table (a sheet with an `ACQUIRER` header,
  or — in QR mode — a `DESTINATION_BANK` / `SOURCE_BANK` / `TRX_DATE` header)
  and skips title rows, repeated page-break headers, and empty rows
- Removes blank/spacer columns (e.g. the empty columns B, D, F in the
  `POS_Transaction_Decline_Report` format)
- Maps common header variants to the mode's canonical schema by **name,
  never by position**:
  - POS Decline: `ACQUIRER, ISSUER, PAN, TRAN_DATE, TIME, TRANS_TYPE,
    AMOUNT, RESP_CODE, Reversal, FE UTRNNO, REFNUM, MERCHANT`
  - POS Success: `ACQUIRER, ISSUER, PAN, TRAN_DATE, TIME, TRANS_TYPE,
    AMOUNT, RESP_CODE, REFNUM, UTRNNO, MERCHANT`
  - POS: `ACQUIRER, ISSUER, CARD_NUMBER, TRANS_DATE, TRANS_TIME,
    TRANS_TYPE, AMOUNT, CURRENCY, RESP, RRN, TERMINAL_ID, ADDRESS`
  - ATM: `ACQUIRER, ISSUER, CARD_NUMBER, TRANS_DATE, TRANS_TIME,
    TRANS_TYPE, AMOUNT, CURRENCY, RESP, RRN, UTRNNO, TERMINAL_ID,
    ADDRESS_NAME`
  - QR: `DESTINATION_BANK, SOURCE_BANK, TRX_DATE, DBTR_ACCT, CDTR_ACCT,
    AMOUNT, TX_ID, STATUS`
- **Handles reshuffled columns**: a report whose columns are in a different
  order (e.g. `UTRNNO` before `RRN`, or `TIME` instead of `TRANS_TIME`)
  still merges into the standard layout. The dashboard flags such files with
  a "reshuffled" badge showing the column order it detected; duplicate
  headers keep the first occurrence
- **Warns about unbalanced columns**: missing columns are reported with the
  column name **and** the file (shown as a banner on the dashboard and in
  the per-file notes); an "unbalanced column counts" warning is shown when
  reports have different numbers of columns
- **Keeps every value exactly as it is in the source file** — literal
  `"null"` strings, raw `TIME`/`TRANS_TIME` values (`115207`, `0:0:20`),
  numbers and their types are never altered
- Sorts the merged result by date, then time (an internal normalized key is
  used for ordering only; the stored values are untouched)
- **Fast for very large reports**: large `.xlsx` files are read with a
  regex-based parser (several times faster than openpyxl, exact same
  values) and written with a manual XML writer (several times faster than
  any general-purpose writer) — the full 428k-row ATM merge completes in
  about a minute
- Dashboard shows summary stats, per-file diagnostics, a response-code chart
  and a preview of the merged table, with one-click download of the merged
  `.xlsx`

## POS, ATM and QR modes

POS, ATM and QR modes merge reports into the SmartVista/daily (plain
tabular) layout:

- sheet name `Report`, **header in row 1** (no title block), plain sheet —
  exactly like the sample daily reports (no fills, borders, column widths
  or frozen panes)
- header variants are matched by name (`TIME` → `TRANS_TIME`,
  `ADDRESS_NAME` → `ADDRESS`, `UTRNNO`/`FE UTRNNO` → `UTRNNO`, `PAN` →
  `CARD_NUMBER`, and in QR mode `Destination Bank` → `DESTINATION_BANK`,
  `Debit Acct` → `DBTR_ACCT`, `Transaction ID` → `TX_ID`, …)
- output filename follows the source convention, e.g.
  `Daily_Tranaction_Report_SmartVista_ATM_15_Aug_26_to_15_Aug_26_Merged.xlsx`
  or `QR_Export_07_Jul_25_to_22_Dec_25_Merged.xlsx`

## Run it

### With uv (recommended — uv is already installed on this machine)

```bash
cd pos-report-merger
uv sync            # creates .venv and installs dependencies
uv run python app.py
```

### With pip

```bash
cd pos-report-merger
python -m venv .venv
.venv\Scripts\activate      # Windows
pip install -r requirements.txt
python app.py
```

Then open <http://127.0.0.1:5000> in your browser.

## Values are never changed

The merger only restructures the table — it never rewrites cell values:

- literal `"null"`/`"nan"` strings stay exactly as they appear (e.g. the
  `null` issuer entries in `POS_Transaction_Decline_Report (42).xls`)
- `TIME` stays in its original format (`115207` stays `115207`, `0:0:20`
  stays `0:0:20`)
- numbers keep their stored value and type (`1216` stays an integer,
  `7184.55` stays a decimal, a text `000001001632` REFNUM stays text)
- sorting uses an internal normalized key purely to order the rows; the
  values written to the workbook are the original ones

## Header matching & missing-column warnings

Columns are matched by **header name**, not position, so column order in the
source files does not matter. Common header variants are recognized, for
example `RESP` / `RESP_CODE`, `FE_UTRNNO` / `FE UTRNNO` / `UTRN NO`,
`ADDRESS_NAME` / `MERCHANT` / `MERCHANT NAME` / `ADDRESS`, `TRANS_DATE` /
`TRAN_DATE` / `TRANS DATE`, `TIME` / `TRANS_TIME`, and `REFNUM` / `REF NUM`.

When a file's column order differs from the standard order, the merger still
maps every column correctly, and the dashboard marks the file with a
"reshuffled" badge showing the detected order. If a header appears twice,
the first column is kept and the duplicate is reported. Unrecognized headers
are dropped and listed in the per-file notes; columns that are recognized
but **not part of the output sample format** are removed and listed as
"removed (not in sample)".

If a file is missing a column that the sample output format requires, a
warning names the missing column and the file — either "these columns are
present in other reports" (when another uploaded file has them) or "these
columns are in the `<sample>` output format".

## How the merged report looks

### POS Decline

```
Row 1:  Report name:  |  POS TRANSACTION  DECLINE REPORT
Row 2:  From Date:    |  20260813
Row 3:  To Date:      |  20260813
Row 4:  ACQUIRER | ISSUER | PAN | TRAN_DATE | TIME | TRANS_TYPE | AMOUNT | RESP_CODE | Reversal | FE UTRNNO | REFNUM | MERCHANT
Row 5+: transaction rows (sorted by date, then time)
```

### POS Success

```
Row 1:  Report name:  |  POS TRANSACTION REPORT
Row 2:  From Date:    |  20260813
Row 3:  To Date:      |  20260813
Row 4:  ACQUIRER | ISSUER | PAN | TRAN_DATE | TIME | TRANS_TYPE | AMOUNT | RESP_CODE | REFNUM | UTRNNO | MERCHANT
Row 5+: transaction rows (sorted by date, then time)
```

### POS / ATM (SmartVista daily layout)

```
Row 1:  ACQUIRER | ISSUER | CARD_NUMBER | TRANS_DATE | TRANS_TIME | TRANS_TYPE | AMOUNT | CURRENCY | RESP | RRN | [UTRNNO] | TERMINAL_ID | ADDRESS[_NAME]
Row 2+: transaction rows (sorted by date, then time)
```

### QR (transfer-export layout)

```
Row 1:  DESTINATION_BANK | SOURCE_BANK | TRX_DATE | DBTR_ACCT | CDTR_ACCT | AMOUNT | TX_ID | STATUS
Row 2+: transaction rows (sorted by date)
```

## Tests

```bash
cd pos-report-merger
uv run pytest
```

The suite includes integration tests against the real report files in the
parent folder (`Aug 13  POS  Declined.xlsx`,
`POS_Transaction_Decline_Report (42).xls`, `Aug 15 ATM s-s .xlsx`,
`Daily_..._ATM_..._1.xlsx`, `Aug 15  POS s-s  .xlsx`,
`Daily_Tranaction_Report_SmartVista_POS_..._xlsx`,
`Aug 13 POS success.xlsx`, `POS_Transaction_SVFE_Report (14).xls`) — they are skipped
automatically when the files are not present.

## Notes on the data

- `PAN` / `CARD_NUMBER` values are already masked in the source reports and
  are kept as-is.
- `REFNUM` values with leading zeros (e.g. `000001001632`) are preserved as
  text; numeric reference numbers are kept numeric.
- Large `FE UTRNNO` / `UTRNNO` numbers keep their exact stored value.
- If one file is missing a column that another file (or the output sample)
  has, the dashboard shows an explicit warning naming the missing column and
  the file; if the reports have different column counts, an "unbalanced
  column counts" warning is shown as well.
- Files that fail to parse are reported per-file on the dashboard and skipped;
  the merge still succeeds as long as at least one file yields transactions.
