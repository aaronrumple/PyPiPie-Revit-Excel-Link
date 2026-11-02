<p align="center"><img src="lib/excellink/excellink.png" width="128" alt="Excel Link"></p>

# Excel Link

Crafted by: [Aaron Rumple, AIA](https://www.linkedin.com/in/aaronrumple/)

A [pyRevit](https://github.com/pyrevitlabs/pyRevit) extension that imports Excel worksheets into Revit drafting views and legends **with visual fidelity**: fonts, colours, borders, fills, merged cells, number formats and alignment. Tables come in at **1:1 printed size** and stay **linked to the workbook** so they can be updated when the spreadsheet changes.

The workbook is read directly; Excel doesn't need to be installed. Part of the PyPiPie tools.

## Features

- **Faithful formatting**: font family, size, bold, italic, underline and colour; theme colours and tints; all 13 Excel border styles; solid, pattern and gradient fills; merged cells; wrap, indent, rotation, shrink-to-fit.
- **Excel number formats**: decimals, thousands separators, %, scientific, fractions, currency and accounting, dates and times, elapsed time, conditional sections, `[Red]`/`[Blue]` colours, K/M scaling, feet-inches (`0"'-"0\"`).
- **True size**: Revit text is sized from each font's cap height so it matches Excel at 1:1; geometry scales with the view so the printed size is always right.
- **Multi-sheet import**: Excel-style sheet tabs with a live preview. Import the visible tab or tick several; each becomes its own drafting view or legend.
- **Linked and updatable**: each table is a detail group carrying its link (Extensible Storage). **Update Links** redraws changed tables in place, even after they've been moved.
- **ACC / Autodesk Docs** through Desktop Connector: links resolve under each user's own `DC` folder, so they work for the whole project team.
- **Open Excel** jumps from a Revit view straight to the linked sheet and range.
- **Light and dark** windows following Revit's UI theme, with the PyPiPie title bar and help (**?** or <kbd>F1</kbd>).

## Buttons

| Button | What it does |
|---|---|
| **Import Excel** | Pick a workbook, preview sheets, choose ranges and options, and import into new drafting views or legends. |
| **Update Links** | Redraw the selected linked tables, or every table whose workbook changed. |
| **Open Excel** | Open the workbook linked to the active view, sheet or selection at the linked sheet and range. |
| **Excel Links** | Manager: status of every link; update, edit (re-point, change sheet/range/options), open, show, unlink. |

When a model opens, the extension notifies you if any links are out of date or their workbook is missing.

## Requirements

| | |
|---|---|
| Revit | Tested on 2024.3. Expected to work on 2021–2026. |
| pyRevit | Tested on 7.0. Expected to work on 4.8 and later. |
| Python engine | IronPython 2.7 (pyRevit's default). |
| Files | `.xlsx`, `.xlsm` |
| Other | None. Excel is optional, only used by **Open Excel**. |

The light/dark window theme needs Revit 2024 or later; older versions use the light theme. On Revit 2025+ (.NET 8), **Open Excel** opens the workbook but can't jump to the linked sheet.

## Installation

### With the pyRevit CLI

```
pyrevit extend ui ExcelLink https://github.com/aaronrumple/PyPiPie-Revit-Excel-Link.git
```

### Manually

1. Download this repository (**Code > Download ZIP**) or clone it.
2. Put the folder in `%APPDATA%\pyRevit\Extensions\` and name it **`ExcelLink.extension`** (the `.extension` suffix is required).
3. In Revit: **pyRevit > Reload**. An **Excel Link** tab appears.

To uninstall, delete the folder and reload pyRevit. Tables already imported stay in your models as ordinary detail groups.

## Quick start

1. **Excel Link > Import Excel**, then **Browse…** for a workbook.
2. Click sheet tabs to preview; tick the ones to import (none ticked = the visible tab).
3. Choose **Drafting view** or **Legend** and a view name pattern (`{sheet}`, `{workbook}`).
4. **Import**. Each sheet opens as a new view at 1:1; place it on a sheet.
5. After editing the workbook, save it in Excel and click **Update Links**.

Try it with `samples/Excel Link Test Workbook.xlsx`, which has one sheet per feature plus a door schedule and finish legend.

Full documentation: [`help/ExcelLink_Help.html`](help/ExcelLink_Help.html), also opened by **?** in the title bar or <kbd>F1</kbd> in either window.

## Limitations

- Conditional formatting (highlight rules, colour scales, data bars, icons) is not imported.
- Strikethrough, charts, pictures, shapes and comments are not imported.
- Mixed fonts, sizes or colours within a single cell use the first run's; bold/italic/underline/super/subscript per run do work.
- Revit can't clip text, so text that Excel cuts off at a neighbouring cell overlaps it in Revit.
- Formula cells show the result last saved by Excel.
- Edits made inside a linked group are replaced on update.

## What it creates

| Element | Naming |
|---|---|
| Detail groups | `Excel - <workbook> - <sheet> - <range>` |
| Text note types | `XL <font> <size>pt [B] [I] [U] #RRGGBB` |
| Filled region types | `XL Fill #RRGGBB` |
| Line styles / patterns | `XL Pen<n> [pattern] #RRGGBB`, `XL Dashed`, `XL Dotted`, … |
| Views | from the name pattern, scale 1:1 |

Link data is stored on each group in Extensible Storage (schema `ExcelLinkData`, GUID `6d3b1f0a-2c8e-4b71-9a55-e1c0b7a4f3d2`, public read) as JSON: workbook path, Desktop Connector location, relative path, sheet, range, options, line weights, and a size/time/SHA-1 fingerprint for change detection.

## Repository layout

```
ExcelLink.tab/Excel.panel/     ribbon buttons (Import Excel, Update Links, Open Excel, Excel Links)
hooks/doc-opened.py            stale-link notification when a model opens
lib/excellink/
    xlsx_reader.py             .xlsx reader: values, styles, theme colours, merges, sizes (stdlib only)
    numfmt.py                  Excel number-format renderer
    layout.py                  worksheet range -> fills, border segments, text in paper points
    revit_build.py             Revit elements and styles from a layout
    link.py                    link storage, change detection, import/update, path resolution
    theme.py, Theme.xaml       light/dark theming following Revit; branded title bar styles
    excellink.png              app logo (window icon, help, README)
    ui.py, *.xaml              import window, manager, commands
help/ExcelLink_Help.html       user guide
samples/                       test workbooks
```

`xlsx_reader`, `numfmt` and `layout` have no Revit dependency and run under CPython 3 for testing.

## License

GPL-3.0. See [LICENSE](LICENSE).

© 2026 [Aaron Rumple, AIA](https://www.linkedin.com/in/aaronrumple/)
