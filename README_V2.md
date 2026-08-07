# Invoice Maker V2

V2 is a parallel upgrade. The original `Module1.bas`, `Sheet1_Invoice.bas`,
`ThisWorkbook.bas`, README, and production template are deliberately unchanged.

The ready-to-use workbook is `Invoice_Template_Formatted_V2.xlsm` in the root
of the invoices share. Its VBA code and buttons are already embedded; the Mac
mini user does not need to open the VBA editor or copy and paste any code.

## What V2 changes

- Saves a non-destructive `.xlsm` copy into the selected customer's
  `Excel`/`EXCEL` folder.
- Exports the matching PDF into that customer's `PDF` folder.
- Finds `Excel` and `PDF` case-insensitively, matching the existing share.
- Avoids Mac Excel's repeated Grant File Access dialogs by creating both files
  in Excel's private temporary directory, then using the installed
  `AppleScriptTask` helper to copy them into the customer's `Excel` and `PDF`
  folders.
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
- Creates large bilingual Save, Print, and New Invoice buttons.
- Places buttons outside the A:E print area and sets their drawing
  `PrintObject` property to false, so they are excluded from paper and PDF.
- Supports multiple printed pages while keeping the invoice one page wide.

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
| `Module1_V2.bas` | Import as a normal VBA module |
| `Sheet1_Invoice_V2.bas` | Copy its event procedure into the **Invoice worksheet** code module |
| `ThisWorkbook_V2.bas` | Copy its event procedures into the real **ThisWorkbook** code module |
| `V2BuildInstaller.bas` | One-time build helper; not needed by invoice users |

Workbook and worksheet event procedures cannot be activated merely by
importing them as ordinary `.bas` modules. They must be placed in the two
document modules shown above.

### Updating `Module1_V2` in the workbook

Editing this file does not change the workbook — the VBA is embedded in
`Invoice_Template_Formatted_V2.xlsm` and must be re-imported.

Close the workbook everywhere first; an SMB lock will block the save. Then in
Excel: **Tools → Macro → Visual Basic Editor**, and

1. **File → Import File…**, press `Cmd-Shift-G`, type the full path to
   `Module1_V2.bas`, and confirm. It arrives as `Module1_V21` because import
   adds a component rather than replacing one.
2. Select the old `Module1_V2` in the project tree, then **File → Remove
   Module1_V2…** and answer **No** to the export prompt.
3. Rename `Module1_V21` back to `Module1_V2` in the Properties pane.
4. **Debug → Compile VBAProject** — it must complete with no dialog.
5. Run the `SaveTemplate` macro. Do not use a normal Save; `Workbook_BeforeSave`
   redirects it into the invoice-save routine and refuses an incomplete invoice.

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

## Mac mini workflow

The Mac deployment scripts (desktop launcher, SMB mount agent, and the
sandbox helper) are site-specific and are not published in this repository.
They live in a local `mac_v2/` folder that is git-ignored, because they
contain the SMB server name and account for a particular installation.

The workflow they provide:

1. Copy the `mac_v2` folder to the target Mac.
2. Open Terminal in that folder and run `zsh install_invoice_shortcut_v2.sh`.
3. Double-click **Invoice Maker V2** on the Desktop.
4. The first time, enter the SMB account password and select **Remember this
   password in my keychain**.

The shortcut connects to `smb://<SMB_USER>@<SERVER>/invoices`, waits for the
share, and opens `Invoice_Template_Formatted_V2.xlsm`. It never contains or
writes the SMB password — macOS Keychain holds it. The SMB username is set by
the `SMB_USER` value in `open-invoice-maker-v2.sh`.

The installer also places `InvoiceMakerV2.applescript` in Excel's required
`~/Library/Application Scripts/com.microsoft.Excel` folder. Do not use `sudo`;
the helper must be installed for the same macOS user who runs Excel.

V2 resolves the share from the workbook's own location first. It also supports
Finder's `/Volumes/invoices` mount and a portable `$HOME/mnt/invoices` mount.

## First functional test

Use a test customer folder and verify all of the following before replacing the
current shortcut:

1. New Invoice clears every old line item after adding at least five items.
2. Canceling the numbered customer prompt leaves the open invoice unchanged.
3. Save creates one `.xlsm` under `Excel` and one `.pdf` under `PDF`.
4. The PDF contains the full invoice but none of the three buttons.
5. A long invoice continues onto page two and remains one page wide.
6. Reboot the Mac mini, use the Desktop shortcut, and confirm that Finder uses
   the Keychain credential to reconnect the SMB share.
7. A `Road Service` line multiplies qty by unit price and reaches the Subtotal
   without appearing in Parts, Labor, or Tire Labor Total.

## Known issues

`EnsureBlankEntryRowV2` inserts a blank entry row above the first `Labor` row.
On Mac Excel the inserted row can inherit the ITEM tag from the row above, so a
stray tag such as `Road Service` appears in column B with no qty, description,
or price. Totals ignore it and the save routine strips empty rows, but a qty
typed on that row would be billed under the inherited type.
