# Invoice Maker V2

V2 is the current production implementation. Its VBA project is intentionally
kept to one standard module plus the two Excel document modules.

The ready-to-use workbook is `Invoice_Template_Formatted_V2.xlsm` in the root
of the invoices share. Its VBA code and buttons are already embedded; the user
on the Mac mini or a Windows PC does not need to open the VBA editor or copy
and paste any code. The same workbook runs on both platforms.

## What V2 changes

- Saves a non-destructive `.xlsm` copy into the selected customer's
  `Excel`/`EXCEL` folder.
- Exports the matching PDF into that customer's `PDF` folder.
- Finds `Excel` and `PDF` case-insensitively, matching the existing share.
- Avoids Mac Excel's repeated Grant File Access dialogs by creating both files
  in Excel's private temporary directory, then using the installed
  `AppleScriptTask` helper to copy them into the customer's `Excel` and `PDF`
  folders.
- Runs on Windows as well, with no helper: saving, printing, and the importer
  use VBA and built-in Windows APIs directly. The Windows file picker is
  filtered to `.xlsm` files.
- Opens the completed Excel invoice automatically after both permanent copies
  succeed, then removes the private temporary staging files.
- Never deletes rows from the open template during a save.
- Rebuilds the complete line-item area during New Invoice, including invoices
  that previously inserted extra rows.
- Keeps the familiar random five-digit invoice number. Existing output files
  are never overwritten if a random number is repeated.
- Prevents filename collisions by adding `-02`, `-03`, and so on rather than
  overwriting an existing invoice.
- Qualifies workbook references, handles partially merged rows, avoids eager
  `IIf` failures, and restores Excel's global event state after errors.
- Locks formulas and labels while leaving intended input cells editable.
- Creates large bilingual Save, Print, and New Invoice buttons, plus an
  Import Invoice button on the production template only.
- The Print button temporarily hides abandoned/empty item rows, opens the
  native Print dialog with two-sided printing disabled, and then restores the
  editable sheet. SAVE INVOICE already creates the archival PDF.
- Blocks File -> Save As from starting the invoice-save workflow a second time;
  users are directed to the Save Invoice or Print buttons instead.
- Repairs the missing Sales Tax summary row and always calculates Total Invoice
  as Subtotal plus Sales Tax, preventing Tire Labor from being counted twice.
- Includes an `ImportLegacyInvoiceV2` importer (the template's Import Invoice
  button) that migrates data into a clean V2 copy, verifies stable totals,
  backs up the original, replaces the selected old XLSM at the same path and
  filename, and replaces its PDF with a freshly rendered one-page copy.
- Places buttons outside the A:E print area and sets their drawing
  `PrintObject` property to false, so they are excluded from paper and PDF.
- Prints on a single page whenever the invoice fits at 75% scale or larger
  (roughly 20 line items). Longer invoices continue onto page two, still one
  page wide. Item rows are one line tall and grow only for a wrapped
  description.
- Paper and PDF output are always black and white (Excel's Black and white
  page setting): fills are dropped and text and borders print black. The
  on-screen invoice keeps its colours.

## Line types

The ITEM column (B) is a dropdown with four values:

| Type | Behaviour |
|---|---|
| `Parts` | Plain `qty x unit price`. Summed into **Parts Total**. |
| `Labor` | Hours roll up into the pinned `Repair Labor @ $80.00/hr` row. Summed into **Labor Total**. |
| `Tires` | Hours roll up into the pinned `Install Tire Labor @ $50.00/hr` row. Summed into **Tire Labor Total**. |
| `Road Service` | Plain `qty x unit price` at whatever rate is entered. No fixed rate, no roll-up, no separate total — it flows straight into the Subtotal. |

`Road Service` is used infrequently and deliberately has no rate bucket: the
unit price is typed per invoice rather than fixed in code. Line items sort
Parts, Labor, Tires, then Road Service.

## V2 source files (developer reference only)

| File | Destination |
|---|---|
| `Module1.bas` | The one standard runtime module |
| `Sheet1_Invoice.bas` | Code for the **Invoice worksheet** document module |
| `ThisWorkbook.bas` | Code for the **ThisWorkbook** document module |

The production VBA project intentionally contains only those three components.
The legacy `Module1` and parallel `Module1_V2` must not coexist: the current
code belongs in the single standard module named `Module1`.

Workbook and worksheet event procedures cannot be activated merely by
importing them as ordinary `.bas` modules. They must be placed in the two
document modules shown above.

### Updating `Module1` in the workbook

Editing this file does not change the workbook — the VBA is embedded in
`Invoice_Template_Formatted_V2.xlsm` and must be re-imported.

Close the workbook everywhere first; an SMB lock will block the save. Then in
Excel: **Tools → Macro → Visual Basic Editor**, and

1. **File → Import File…**, press `Cmd-Shift-G`, type the full path to
`Module1.bas`, and confirm. Remove the existing standard `Module1` first so
the imported component keeps the correct name.
2. Copy the short event procedures from `Sheet1_Invoice.bas` and
`ThisWorkbook.bas` into their matching document modules.
3. Confirm the project tree contains only `ThisWorkbook`, `Sheet1`, and
`Module1`.
4. **Debug → Compile VBAProject** — it must complete with no dialog.
5. Run `InitializeInvoiceMakerV2` once so summary rows and formatting are
   repaired, then run `SaveTemplateV2`. Do not use a normal Save; `Workbook_BeforeSave`
   redirects it into the invoice-save routine and refuses an incomplete invoice.

On Windows the steps are the same with Windows shortcuts. Close the workbook
everywhere, then press **Alt+F11**:

1. Right-click the existing `Module1` and choose **Remove Module1** (do not
   export), then **File → Import File…** and pick `Module1.bas`.
2. Confirm the project tree contains only `ThisWorkbook`, `Sheet1`, and
   `Module1`.
3. **Debug → Compile VBAProject** — it must complete with no dialog.
4. Run `InitializeInvoiceMakerV2`, then `SaveTemplateV2`.

### Upgrading one old invoice

On the production template, click **IMPORTUJ FAKTURĘ / IMPORT INVOICE**, then
select the old `.xlsm`. The macro list also offers it as
`ImportLegacyInvoiceV2` (**Control+Option+F8** on the Mac, **Alt+F8** on
Windows). It works on V1 invoices and on earlier V2 saves alike, on both
platforms. When Mac Excel asks while the importer opens its read-only and
staged copies, choose **Disable Macros**. Review the old/new total comparison
and approve the replacement only if it is correct. The invoice number, invoice
date, ship/due dates, data, path, and filename are retained. A hash-verified
copy of the original is kept in an `Invoice Import Backups` folder on the
Desktop: `~/Desktop` on the Mac, and the real Windows Desktop (including a
OneDrive-redirected one) on Windows.

The importer also renders a new one-page PDF and puts it in the customer's
`PDF` folder under the invoice's own name (beside the XLSM when there is no
`Excel`/`PDF` folder pair). An existing PDF of that name is backed up to the
same Desktop folder first. The PDF is rendered before anything is replaced,
so a rendering failure leaves the original untouched; a failure while placing
it is reported as a warning, and running the import again retries it.

On the Mac the file work goes through the AppleScript helper. On Windows it is
native and needs no helper. The importer stages in `%TEMP%\InvoiceMakerV2`,
fingerprints files with SHA-256 through the built-in Windows CNG API
(`bcrypt.dll`), verifies every backup and copy against those hashes, and
swaps the new file in with a same-folder rename (`MoveFileExW`, replacing the
existing file). If the swapped file does not verify, the original is restored
from the verified backup. File access uses `FileSystemObject` and
`ADODB.Stream`, so customer folder names with Polish letters work.

For developers: a private `RunUnattendedImportV2(legacyPath, backupRoot)`
entry point runs the same import without dialogs, writing backups under
`backupRoot`, so it can be tested end to end on Windows through
`Application.Run`. It is not exposed as a macro.

Keep `Module1.bas` pure ASCII. The VBA editor imports `.bas` files in the
legacy Mac Roman / Windows-1252 code page, not UTF-8, so a Polish letter typed
directly in the source turns into junk characters (the Save button once read
`ZAPISZ FAKTURƒò`). Write such letters as backslash-u escapes with four hex
digits inside a string passed to `UnicodeTextV2`, as the button captions do.

Do not paste the module with `Cmd-A`/`Cmd-V` through UI automation. If the code
pane does not hold keyboard focus, macOS drops the Command modifier and types a
literal `av` before `Option Explicit`, which fails to compile with "Invalid
outside procedure". Importing avoids the clipboard entirely.

Verify afterwards by reading the dropdown back rather than trusting a clean
compile — a compile succeeds against unchanged code:

```
osascript -e 'tell application "Microsoft Excel" to return formula1 of ¬
  validation of range "B15" of worksheet "Invoice" of workbook ¬
  "Invoice_Template_Formatted_V2.xlsm"'
```

## Mac deployment requirements

Deployment is site-specific and is not published in this repository, because it
encodes the SMB server name and account for a particular installation. What
follows is what any such setup has to provide.

Everything in this section applies to the Mac only. Windows needs nothing
beyond the workbook itself and access to the invoices share.

**The sandbox helper is required.** Mac Excel is sandboxed and cannot write to
SMB folders without repeatedly prompting for file access. V2 works around this
by writing to Excel's private temporary directory and delegating the copy to an
AppleScript helper running outside the sandbox.

The helper must be installed at:

```
~/Library/Application Scripts/com.microsoft.Excel/InvoiceMakerV2.applescript
```

That filename and location are fixed — `Module1.bas` calls
`AppleScriptTask("InvoiceMakerV2.applescript", ...)`. Install it for the same
macOS account that runs Excel, without `sudo`; the folder is per-user and a
root-owned copy will not be found.

The helper lists customer folders, moves staged invoice/PDF pairs, controls
temporary one-sided printer defaults, performs verified legacy replacements,
and places the re-rendered PDF of an imported invoice (`placeImportedPdf`).
Reinstall the helper whenever `Module1` gains a new `AppleScriptTask` call;
an older helper makes the importer report that the PDF was not updated.
Copy files without preserving permission bits — SMB shares commonly reject
`fchmod`, which fails the copy after the data has already been written.

**Share access.** The user needs the `invoices` SMB share mounted, with the
password stored in the macOS Keychain rather than in any script. A desktop
launcher that mounts the share and opens
`Invoice_Template_Formatted_V2.xlsm` gives a non-technical user a single icon
to click.

V2 resolves the share from the workbook's own location first. It also supports
Finder's `/Volumes/invoices` mount and a portable `$HOME/mnt/invoices` mount.

## First functional test

Use a test customer folder and verify all of the following before replacing the
current shortcut:

1. New Invoice clears every old line item after adding at least five items.
2. Canceling the numbered customer prompt leaves the open invoice unchanged.
3. Save creates one `.xlsm` under `Excel` and one `.pdf` under `PDF`.
4. The PDF contains the full invoice but none of the three buttons.
5. An invoice of up to about 20 lines prints on one page; a longer one
   continues onto page two and remains one page wide.
6. Reboot the Mac mini, use the Desktop shortcut, and confirm that Finder uses
   the Keychain credential to reconnect the SMB share.
7. A `Road Service` line multiplies qty by unit price and reaches the Subtotal
   without appearing in Parts, Labor, or Tire Labor Total.
8. Paper and PDF output are black and white.
9. The button captions show Polish letters correctly (`ZAPISZ FAKTURĘ`).
10. Import Invoice on the template upgrades a test invoice and replaces its
    PDF with a one-page copy, backing up both originals.
11. On Windows, open the same template from the share and repeat Save, Print,
    and Import Invoice on a test customer folder.

## Known issues

`EnsureBlankEntryRowV2` inserts a blank entry row above the first `Labor` row.
On Mac Excel the inserted row can inherit the ITEM tag from the row above, so a
stray tag such as `Road Service` appears in column B with no qty, description,
or price. Totals ignore it and the save routine strips empty rows, but a qty
typed on that row would be billed under the inherited type.
