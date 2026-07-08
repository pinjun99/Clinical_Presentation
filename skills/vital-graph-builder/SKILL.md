---
name: vital-graph-builder
description: Build native Excel vital-sign charts from the graph-selection sheets produced by vitals-chart-extractor. Use when the user has ticked Graph? checkboxes on the BP, Heart Rate, Respiratory, SpO2, or Pain Score sheets and wants per-vital charts generated or regenerated - a blood pressure chart with vertical double-arrow lines between systolic and diastolic plus a data table, and single-colour line charts with data tables for the other vitals.
---

# Vital Graph Builder

Use this skill to turn ticked rows on the vital graph-selection sheets (created by `vitals-chart-extractor`) into native Excel charts. Everything for one vital lives on that vital's own sheet: the tick table in columns A-F, and the horizontal data block plus chart from column H.

## How to run

```bash
python scripts/build_vital_graphs.py [workbook.xlsx]
```

The workbook path defaults to `outputs/vitals_extract/vitals_extract.xlsx`. The script:

1. Backs up the workbook first.
2. Reads rows where the `Graph?` checkbox is ticked (`TRUE`) on each of `BP`, `Heart Rate`, `Respiratory`, `SpO2`, `Pain Score`.
3. Re-lists the ticked readings horizontally (column H onward, static resolved values) on the same sheet.
4. Draws one chart per vital anchored at H5.
5. Post-patches the saved .xlsx (see "Why the post-save patch exists").

Re-run it whenever the user changes their ticks; it is idempotent and rebuilds all five charts.

## Chart appearance (do not change without user request)

- **BP**: vertical double-arrow lines from Diastolic up to Systolic, dark grey (#404040) 1.5pt with triangle arrowheads at both ends, built as custom minus-direction error bars on a hidden Systolic series (arrow lengths come from a `Sys-Dia (arrow length)` helper row). Data table below shows Systolic and Diastolic only. Y-axis visible, minimum 50.
- **Heart Rate / Respiratory / SpO2 / Pain Score**: one standard blue (#4472C4) line with circle markers, data table below, visible y-axis.
- All charts: units in the chart title (for example `Heart Rate (bpm)`), no y-axis title, no legend, title above the plot area, horizontal gridlines on.

## Why the post-save patch exists (hard-won; do not "simplify" away)

openpyxl alone cannot produce the required output. The script string-patches the saved .xlsx package because, verified against real Excel rendering:

- **Checkboxes**: Excel-native checkbox cells are booleans plus a `FeaturePropertyBag` part; openpyxl strips it on save, degrading ticks to TRUE/FALSE text. The patch re-injects the part, content type, workbook relationship, and checkbox cell format so column A stays a real tick column.
- **Axes**: openpyxl emits bare-bones axis XML that Excel renders without labels. The patch replaces `catAx`/`valAx` with complete axis XML (delete=0, tickLblPos=nextTo, numFmt, spPr, txPr). Visible axes and the chart data table coexist fine.
- **Title position**: without `<overlay val="0"/>` Excel floats the title over the plot. The patch adds it, plus `<autoTitleDeleted val="0"/>`.

Other constraints baked into the script:

- Excel `hiLowLines` and stock charts do NOT render in Mac Excel at all - not even hand-written XML. Custom error bars are the only construct that draws the BP arrows. Never switch BP back to high-low lines.
- The horizontal data block must hold static resolved values; formula-linked chart source cells stop Excel drawing the data table and axis labels.
- Set an explicit colour on both the line and its markers, otherwise Excel theme-cycles marker colours.

## Verifying output through real Excel (macOS)

Never judge chart correctness from openpyxl or a preview library; render through Excel itself:

1. Copy the workbook into a `/private/tmp` scratch folder (Excel's sandbox only writes PDFs next to a file it opened there; writes elsewhere silently no-op).
2. AppleScript: `open (POSIX file ...)` then `save wb in (POSIX file ".../x.pdf") as PDF file format`, wrapped in `with timeout of 240 seconds`.
3. Convert pages with pymupdf and inspect the chart pages (locate them by `page.get_drawings()` count).

Caveats: whole-workbook PDF export truncates around 61 pages, cutting late sheets - to verify a late sheet, delete the other worksheets in a throwaway copy via AppleScript first. Never set a print area; it drops axis numbers and the data table from the export only.

## Safety

The charts visualise transcribed clinical data. Remind the user that clinically important readings should be verified against the original chart or by appropriate clinical staff.
