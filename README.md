# excel-invoice-macros

Excel VBA macros for an invoice template used in a small trucking and fleet
maintenance business. The template handles save, reset, print layout, and
automatic calculation of line items, taxes, and totals.

Two generations live here side by side.

| | [`v1/`](v1/) | [`v2/`](v2/) |
|---|---|---|
| Status | Stable, superseded | Current |
| Platform | Windows and Mac | Mac (Apple Silicon, SMB share) |
| Audience | Developer imports the modules by hand | Non-technical user clicks a button |
| Install | Import three `.bas` files into a workbook | VBA already embedded in the workbook |
| Saving | Prompts for a folder, writes one `.xlsm` | Writes `.xlsm` and PDF into per-customer folders |

**V1 is kept deliberately.** It is a known-good copy that predates the V2 work,
and it remains the reference for the original Windows-compatible behaviour.
Nothing in V2 modifies it.

## Which one do I want?

Use **V2** if you are running the invoice program. It is what is deployed on
the Mac mini: the workbook opens from a desktop shortcut, the VBA and buttons
are already inside it, and saving files into the right customer folders is
automatic.

Use **V1** if you want the original, smaller module set, need Windows support,
or want to see the behaviour V2 was built from.

## What V2 changed

V2 was a reliability and usability pass over V1, not a rewrite. The main
differences:

- **Non-destructive saving.** V1 deleted blank and `$0` rows *before* asking
  where to save, so cancelling the dialog left the invoice damaged. V2 stages
  everything and never deletes rows from the open template.
- **Merge-safe reset.** `Range.ClearContents` raised error 1004 on merged rows
  in some Mac Excel builds. V2 clears cell by cell.
- **No sandbox prompts.** Mac Excel asked for "Grant File Access" repeatedly on
  SMB paths. V2 writes to Excel's private temp directory and hands the files to
  an `AppleScriptTask` helper outside the sandbox.
- **Error recovery.** V1 could leave `EnableEvents` off after a failure, making
  Excel appear permanently broken. V2 restores global state on every path.
- **Buttons.** Large bilingual Save, Print, and New Invoice buttons, excluded
  from both paper and PDF output.
- **Road Service line type**, alongside Parts, Labor, and Tires.

Full details in [`v2/README.md`](v2/README.md).

## Line types

The ITEM column (B) drives sorting and the summary totals.

| Type | Behaviour | In V1? |
|---|---|---|
| `Parts` | Plain `qty x unit price` → **Parts Total** | yes |
| `Labor` | Hours roll into the `$80.00/hr` repair row → **Labor Total** | yes |
| `Install` | Normalised to `Labor` | yes |
| `Tires` | Hours roll into the `$50.00/hr` tire row → **Tire Labor Total** | yes |
| `Road Service` | Plain `qty x unit price`, no fixed rate, no roll-up, no separate total — flows straight into the Subtotal | V2 only |

## Repository layout

```
v1/   original modules and their README
v2/   current modules, build helper, and their README
```

The Mac deployment pieces (desktop launcher, SMB mount agent, sandbox helper)
are **not** published here — they encode the SMB server name and account for a
specific installation. [`v2/README.md`](v2/README.md) documents what a
deployment has to provide, including the sandbox helper that V2 requires.

The `.xlsm` workbooks are also git-ignored; the VBA source in this repository
is the authoritative copy.

## License

MIT
