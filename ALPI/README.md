# Monthly Summary by Physician (VBA / Excel)

A VBA macro that starts from a **database exported from the management system**, **aggregates amounts by individual physician**, and generates a **monthly summary report**.

> ⚠️ This repository contains **code only**. It does not include real data: no personal records, amounts, physician names, patient information, or email addresses.

---

## What It Does

1. Imports the raw data export from the management system (one row per medical service).
2. Creates a **working copy** and adds four calculated columns:

   | Column | Description |
   |---|---|
   | `Costo corretto` | Amount with the correct sign: **reversals** are subtracted |
   | `Bollo` | €2 charged only once per invoice, if the invoice total exceeds €77.47 (in absolute value) |
   | `Chiave` | Physician's name extracted from the agenda description (text before the opening parenthesis) |
   | `Matricola` | Physician's employee ID, retrieved from the master data using the key |

3. Builds a **pivot table by employee ID** (sum of corrected costs and stamp duty).
4. Transfers the totals to the **Riepilogo** worksheet, with one row per physician. Physicians with no transactions during the period are left blank.

```mermaid
flowchart LR
    A[Management System Export] --> B[Working Copy]
    B --> C[Corrected Cost, Stamp Duty, Key, Employee ID]
    D[(Physician Master Data)] --> C
    C --> E[Pivot by Employee ID]
    E --> F[Summary Worksheet]
```

---

## Repository Contents

| File | Description |
|---|---|
| `RiepilogoALPI.bas` | VBA module containing the macros |
| `README.md` | This document |

### Included Macros

| Macro | Function |
|---|---|
| `Riepilogo` | Executes the entire process described above |
| `Eliminare_valori_negativi` | Utility: converts formulas in the selected range into values and replaces every number with its absolute value (cannot be undone) |

---

## Expected Workbook Structure

Worksheet names are defined as constants at the beginning of the module and can be changed in one place only.

| Worksheet | Contents |
|---|---|
| `estrazione LP4` | Management system export, with headers in row 1. Required columns: `storno`, `importo`, `numero_fattura`, `desc_agenda`. The data must be **sorted by invoice number** |
| `DB Medici` | Physician master data with 4 columns: `Chiave`, `Matricola`, `Cognome`, `Nome` |
| `Riepilogo` | Final summary report: employee ID in column A, data starting from row 4 (row 3 contains headers). The macro populates column D (online amount) and column F (online stamp duty) |

`DB Risultato` and `Pvt` (pivot table) are **recreated every time the macro runs**. Do not edit them manually.

Example of fictional `DB Medici` data:

| Chiave | Matricola | Cognome | Nome |
|---|---|---|---|
| ROSSI MARIO | 1001 | ROSSI | MARIO |
| BIANCHI LUISA | 1002 | BIANCHI | LUISA |

In the data export, `desc_agenda` must follow this format: `ROSSI MARIO (agenda description)`. The key is the text preceding ` (`.

---

## How to Use It

1. Open an Excel workbook containing the worksheets described above and save it as an `.xlsm` file.
2. Press `Alt + F11` to open the VBA editor.
3. Select `File → Import File…` and choose `RiepilogoALPI.bas`.
4. Paste the new data export into the `estrazione LP4` worksheet.
5. Run `Riepilogo` (`Alt + F8`, or assign it to a button).
6. Check the `Riepilogo` worksheet.

---

## Technical Design Choices

- Export columns are identified **by header name**, not by column letter. The macro therefore remains functional if the management system adds or moves columns.
- Formulas are written in **R1C1 notation across the entire range** in a single instruction, without using `Select` or `FillDown`.
- The stamp duty threshold is stored as text (`"77.47"`) because R1C1 formulas require a decimal point even when using the Italian version of Excel.
- Physicians with no transactions are identified using `IsError`, without depending on Excel's language settings (`#N/D` / `#N/A`).
- The pivot table uses only the four calculated columns, which have unique headers.

---

## Known Limitations

- If `desc_agenda` does not contain an opening parenthesis, the `Chiave` formula returns an error and the row is not assigned to any physician.
- Stamp duty calculation assumes that the export is sorted by invoice number.
- The original `Riepilogo` workbook also includes manually populated columns (payments and deductions); these are not handled by the macro.

---

## Privacy

- Never upload `.xlsm`/`.xlsx` files, real data exports, or physician master data to GitHub. Excel files may also retain hidden data (pivot table caches, hidden worksheets, metadata).
- Any example file added to the repository must contain fictional data only.
- Suggested `.gitignore`:

```gitignore
*.xlsx
*.xlsm
*.xls
*.csv
~$*
```

---

## Requirements

- Microsoft Excel for Windows (with VBA enabled)
