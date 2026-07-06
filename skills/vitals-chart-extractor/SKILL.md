---
name: vitals-chart-extractor
description: Extract handwritten or image-based clinical vitals observation charts into review-first Excel workbooks. Use when Codex needs to read clinical observation chart images, scans, or provided chart captures containing date, time, BP, pulse, respiratory rate, SpO2, pain score, or similar vitals, create a spreadsheet, highlight uncertain readings, apply date/time sequence sense, support user review, and produce a final sorted workbook.
---

# Vitals Chart Extractor

Use this skill to convert handwritten clinical vitals or observation chart images into an Excel workbook that is easy for a human to review first, then sort after confirmation, then select rows for later graphing.

## Core Principle

Preserve source order before sorting. The first workbook must follow the source image or chart row order so the user can compare Excel rows against the original chart. Sort chronologically only after the user reviews and confirms highlighted cells.

## Output Structure

Create a workbook with these sheets when possible:

- `Review - Source Order`: rows in the same order as the provided chart images or source material.
- `Final - Sorted`: generated only after user confirmation; sorted by Date then Time.
- Vital selection sheets after final confirmation: `BP`, `Heart Rate`, `Respiratory`, `SpO2`, and `Pain Score`.
- `Notes`: legend, assumptions, source files, and review status.

For small or iterative jobs, a single `Vitals Extract` sheet is acceptable, but keep the same review-first workflow.

## Columns

Use these columns unless the user requests otherwise:

- Source File
- Photo Row Number
- Date
- Time
- BP (mmHg)
- Systolic
- Diastolic
- Pulse (bpm)
- Respiratory (breaths/min)
- SpO2 (%)
- Pain Score (0-10)
- Notes

Keep BP as written, and fill Systolic/Diastolic only when both numbers are clear. Use `AM`/`PM` consistently for Time. Store pulse, respiratory, SpO2, and pain score as plain numbers where possible; put units in headers, not in cell values.

## Final Graph Selection Sheets

After the user confirms review and the sorted/final data is ready, add separate vital-specific sheets in the same workbook for graph selection.

Use these columns:

- `BP`: Graph?, Date, Time, BP (mmHg), Systolic, Diastolic
- `Heart Rate`: Graph?, Date, Time, Pulse (bpm)
- `Respiratory`: Graph?, Date, Time, Respiratory (breaths/min)
- `SpO2`: Graph?, Date, Time, SpO2 (%)
- `Pain Score`: Graph?, Date, Time, Pain Score (0-10)

Use real clickable Excel checkbox cells in `Graph?` when the workbook-writing library supports them. The checkbox value must be readable as `TRUE`/`FALSE` later; Skill 2 will graph only rows where `Graph?` is `TRUE`. If native checkbox cells are not supported, clearly tell the user and use a simple readable fallback such as `TRUE`/`FALSE`.

Date, Time, and vital values should be formula-linked to the final sorted/source sheet where possible so later edits to the confirmed data update the selection sheets automatically. Do not include `Source File` or `Notes` on these graph-selection sheets once the data is final; keep traceability columns only in the main vitals sheet.

Keep normal Excel gridlines visible. Users need row lines for checking and ticking boxes.

Keep the `Notes` sheet short after final confirmation: include only the color legend if still relevant, checkbox meaning, Skill 2 read rule, and a brief clinical caution.

## Color Rules

Color only the specific cell that needs review, not the whole row.

- No color: confirmed/no issue.
- Blue: date/time assumption, date/time sequence issue, blank date, invalid time, or inferred date/time.
- Orange: uncertain clinical or vital value, such as BP, pulse, respiratory rate, SpO2, or pain score.

Add a concise explanation in Notes whenever a cell is colored.

## Date And Time Sense

Use common sequence sense for Date and Time only. The goal is to reduce the user's review work while still showing assumptions clearly.

If a time appears impossible in row order, make the most reasonable date/time assumption and color the assumed cell blue. Example: if source order shows `7AM -> 10AM -> 2PM -> 10PM -> 4PM`, and the final handwritten time could plausibly be `11PM`, enter `11PM`, color that Time cell blue, and note: `Assumed 11PM because original looked like 4PM but appears after 10PM in source order.`

Also color blue when:

- time goes backward within the same date in source order;
- date is missing or ambiguous;
- date jumps unexpectedly compared with surrounding rows;
- a date/time is inferred from nearby rows.

Do not use date/time sequence assumptions for clinical values.

## Clinical Values

Do not predict or clinically normalize BP, pulse, respiratory rate, SpO2, or pain score. If a value is unclear, transcribe the best visible reading, color that exact cell orange, and explain the uncertainty in Notes.

Examples:

- `SpO2 unclear; appears like 96 or 98.`
- `Diastolic unclear.`
- `Pulse appears amended.`

Never change a clinical value simply because it seems medically more likely.

## Workflow

1. Identify all provided chart images/scans and convert unsupported formats such as HEIC to viewable copies if needed. Do not alter originals.
2. Read each chart image in visible row order. Preserve source order and source file for every row.
3. Build the `Review - Source Order` sheet.
4. Apply the color rules: blue for date/time assumptions or sequence concerns, orange for uncertain clinical values.
5. Include a `Notes` sheet with the color legend, source list, and caution that handwritten clinical data needs human confirmation.
6. Ask the user to review colored cells. Do not create the final sorted sheet until the user confirms review is done.
7. After user confirmation, make a backup of the user-edited workbook before changing it.
8. Treat the user-edited workbook as the source of truth.
9. Re-check date/time sequence. If issues remain, highlight them blue and report them.
10. Create `Final - Sorted` sorted by Date then Time, or update the final workbook if the user explicitly wants a single sheet.
11. Create the five vital-specific graph-selection sheets with readable native checkboxes in `Graph?`.
12. Verify Date cells display as human-readable dates such as `d/m`, not Excel serial numbers like `46175`.
13. Verify formulas/errors/rendering before final delivery.

## Review Checks

Before final delivery, check:

- no unsupported image files remain unread;
- all extracted rows have source file and row number;
- time format is consistent;
- date display is human readable;
- blue/orange cells are cell-specific;
- Notes explains each assumption or uncertainty;
- final sorted sheet has no backward time jumps within the same date unless intentionally flagged.
- graph-selection sheets contain only Date, Time, the relevant vital value(s), and a readable checkbox column;
- checkbox cells can be read later as `TRUE`/`FALSE`;
- normal gridlines are visible unless the user explicitly asks for a print-style worksheet.

## Safety

This skill produces a transcription aid, not a clinical authority. Tell the user that uncertain or clinically important readings should be verified against the original chart or by appropriate clinical staff.
