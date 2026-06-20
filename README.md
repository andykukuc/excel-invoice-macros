# excel-invoice-macros

Excel VBA macros for an invoice template used in a small trucking/fleet maintenance business. The template runs on both Windows and Mac (Samba share), handles save, reset, PDF export, print layout, and auto-calculation of line items, taxes, and totals.

## Modules

| File | Description |
|------|-------------|
| `Module1.bas` | Core logic — save/export, reset, sort, formatting, line item helpers |
| `ThisWorkbook.bas` | Workbook event handlers — triggers sort, formatting, and save flow on BeforeSave |
| `Sheet1_Invoice.bas` | Worksheet change handler — auto-updates line amounts as data is entered |

## Features

- **SaveInvoice** — builds a filename from fleet number, model, invoice number, and date; saves a clean `.xlsm` copy (no auto PDF); deletes `$0` rate-bucket rows, blank rows, and Tire Labor Total if no tires before saving; opens the saved copy and closes the template
- **ExportPDF** — exports the invoice as a letter-size portrait PDF with 0.5" margins; deletes empty rows before export for a clean printout (not called automatically — user prints to Adobe PDF printer manually)
- **SortLineItems** — sorts line items in memory (Parts → Labor → Tires); rate-bucket rows pinned to the bottom of their section; called from `FitToOnePage`
- **UpdateLineAmounts** — rolls `Labor`/`Install`-tagged hours into the `$80.00/hr` repair row and `Tires`-tagged hours into the `$50.00/hr` tire row; recalculates all amount columns
- **UpdateFormulas** — rebuilds Subtotal, Parts Total, Labor Total, Tire Labor Total, Sales Tax, Total Invoice, and Total Due using explicit `SUMIF` tags; inserts missing summary rows automatically
- **FormatInvoice** — applies alternating row shading, currency formatting, `0.##` QTY format (so half-hours display correctly), summary section row height, and Parts/Labor/Tires dropdown on the ITEM column; does **not** touch cell borders so template borders are preserved
- **HideEmptyBuckets** — hides the `$80/hr` repair row + Labor Total and `$50/hr` tire row + Tire Labor Total when their amounts are `$0`; used by `FitToOnePage` for print preview
- **ShowAllBuckets** — un-hides all rows from the line item block through Total Due; called automatically when editing starts so no rows are stuck hidden
- **AlignLineItems** — left-aligns and wraps text in the line item description column
- **ResetTemplate** — clears all invoice fields, generates a new invoice number, and re-applies formatting
- **FitToOnePage** — sorts line items, refreshes calculations, hides empty buckets, sets print area with 0.5" margins fit to one page tall, and opens print preview
- **Auto-fill (Sheet1)** — typing in A7 (Bill To) auto-fills A8, C7, C8 (Ship To mirrors Bill To); typing in A11 auto-fills B11
- Works on **Mac and Windows** — Mac uses `SaveCopyAs` to avoid SMB atomic-write issues; Windows uses `GetSaveAsFilename`

## Setup

1. Open your `.xlsm` invoice template in Excel
2. Open the VBA editor (`Alt+F11` on Windows, `Cmd+Option+F11` on Mac)
3. Import each `.bas` file: **File → Import File**
4. Assign macros to buttons on the sheet as needed:
   - `SaveInvoice` → Save button
   - `ResetTemplateManual` → New Invoice button
   - `FitToOnePage` → Print Preview button

## Line Item Tagging

The `B` column (tag) drives sorting and summary totals:

| Tag | Sorts into | Rolls into |
|-----|-----------|------------|
| `Parts` | Parts section | Parts Total |
| `Labor` | Labor section | Labor Total (hours → $80/hr repair row) |
| `Install` | Labor section | Labor Total (normalized to `Labor`, hours → $80/hr repair row) |
| `Tires` | Tires section | Tire Labor Total (hours → $50/hr tire row) |

## Mac Notes

- The Samba share must be mounted at `/Users/andykukuc/mnt/invoices/` (edit `baseDir` in `Module1.bas` to match your mount point)
- Uses `InputBox` for folder selection since Mac Excel can't use `GetSaveAsFilename` with file filters on network paths
- Uses `SaveCopyAs` instead of `SaveAs` to avoid the repair dialog on SMB shares

## License

MIT
