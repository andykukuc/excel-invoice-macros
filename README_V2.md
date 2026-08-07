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
