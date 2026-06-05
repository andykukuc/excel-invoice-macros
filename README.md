# excel-invoice-macros

Excel VBA macros for an invoice template used in a small trucking/fleet maintenance business. The template runs on both Windows and Mac (Samba share), handles save, reset, PDF export, print layout, and auto-calculation of line items, taxes, and totals.

## Modules

| File | Description |
|------|-------------|
| `Module1.bas` | Core logic — save/export, reset, formatting, line item helpers |
| `ThisWorkbook.bas` | Workbook event handlers — triggers formatting and save flow on BeforeSave |
| `Sheet1_Invoice.bas` | Worksheet change handler — auto-updates line amounts as data is entered |

## Features

- **SaveInvoice** — builds a filename from fleet number, model, invoice number, and date; exports both an `.xlsm` copy and a `.pdf` to your chosen folder
- **ExportPDF** — exports the invoice as a letter-size portrait PDF; flows to page 2 if needed instead of forcing everything onto one page
- **FormatInvoice** — applies alternating row shading, light borders, and currency formatting to the line item block
- **UpdateLineAmounts** — auto-rolls hours tagged `Labor` into the `$80.00/hr` repair row and hours tagged `Tires` into the `$50.00/hr` install row; recalculates all amount columns
- **UpdateFormulas** — rebuilds Subtotal, Parts Total, Labor Total, Sales Tax, Total Invoice, and Total Due formulas using explicit `SUMIF` tags; inserts a Labor Total row if missing
- **AlignLineItems** — left-aligns and wraps text in the line item description column
- **ResetTemplate** — clears all invoice fields, generates a new invoice number, and re-applies formatting
- **FitToOnePage** — refreshes calculations and formatting, sets print area, and opens print preview
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

The `B` column (tag) drives the summary totals:

| Tag | Rolls into |
|-----|-----------|
| `Parts` | Parts Total |
| `Labor` | Labor Total (also guards the $80/hr and $50/hr hourly rows) |
| `Tires` | rolled into the $50/hr tire install labor row |

## Mac Notes

- The Samba share must be mounted at `/Users/andykukuc/mnt/invoices/` (edit `baseDir` in `Module1.bas` to match your mount point)
- Uses `InputBox` for folder selection since Mac Excel can't use `GetSaveAsFilename` with file filters on network paths
- Uses `SaveCopyAs` instead of `SaveAs` to avoid the repair dialog on SMB shares

## License

MIT
