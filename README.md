# excel-invoice-macros

Excel VBA macros for an invoice template used in a small trucking/fleet maintenance business. The template runs on both Windows and Mac (Samba share), handles save, reset, print layout, and auto-calculation of line items, taxes, and totals.

## Modules

| File | Description |
|------|-------------|
| `Module1.bas` | Core logic — save invoice, reset template, fit-to-page print setup, line item helpers |
| `ThisWorkbook.bas` | Workbook event handlers — triggers formatting and save flow on BeforeSave |
| `Sheet1_Invoice.bas` | Worksheet change handler — auto-updates line amounts as data is entered |

## Features

- **SaveInvoice** — builds a filename from fleet number, model, invoice number, and date; lets you pick a folder on the Samba share via InputBox; saves a copy without touching the template
- **ResetTemplate** — clears all invoice fields and generates a new random invoice number
- **FitToOnePage** — hides empty line item rows, sets print area, and opens print preview scaled to one page
- **AlignLineItems / UpdateLineAmounts / UpdateFormulas** — keeps subtotal, tax, and total due formulas in sync as line items change
- Works on **Mac and Windows** — Mac uses `SaveCopyAs` to avoid SMB atomic-write issues

## Setup

1. Open your `.xlsm` invoice template in Excel
2. Open the VBA editor (`Alt+F11` on Windows, `Cmd+Option+F11` on Mac)
3. Import each `.bas` file: **File → Import File**
4. Assign macros to buttons on the sheet as needed:
   - `SaveInvoice` → Save button
   - `ResetTemplateManual` → New Invoice button
   - `FitToOnePage` → Print Preview button

## Mac Notes

- The Samba share must be mounted at `/Users/andykukuc/mnt/invoices/` (edit `baseDir` in `Module1.bas` to match your mount point)
- Uses `InputBox` for folder selection since Mac Excel can't use `GetSaveAsFilename` with file filters on network paths
- Uses `SaveCopyAs` instead of `SaveAs` to avoid the repair dialog on SMB shares

## License

MIT
