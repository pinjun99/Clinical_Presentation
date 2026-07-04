---
name: vitals-chart-extractor
description: Extract handwritten or photographed hospital vitals observation charts into review-first Excel workbooks. Use when Codex needs to read clinical observation chart photos or scans containing date, time, BP, pulse, respiratory rate, SpO2, pain score, or similar vitals, create a spreadsheet, highlight uncertain readings, apply date/time sequence sense, support user review, and produce a final sorted workbook.
---

# Vitals Chart Extractor

Use this skill to convert handwritten hospital vitals or observation chart photos into an Excel workbook that is easy for a human to review first, then sort after confirmation.

## Core Principle

Preserve source order before sorting. The first workbook must follow the photo or chart row order so the user can compare Excel rows against the original chart. Sort chronologically only after the user reviews and confirms highlighted cells.

## Output Structure

Create a workbook with these sheets when possible:

- `Review - Photo Order`: rows in the same order as the source photos/chart.
- `Final - Sorted`: generated only after user confirmation; sorted by Date then Time.
- `Notes`: legend, assumptions, source files, and review status.

For small or iterative jobs, a single `Vitals Extract` sheet is acceptable, but keep the same review-first workflow.

## Columns

Use these columns unless the user requests otherwise:

- Source Photo
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

## Color Rules

Color only the specific cell that needs review, not the whole row.

- No color: confirmed/no issue.
- Blue: date/time assumption, date/time sequence issue, blank date, invalid time, or inferred date/time.
- Orange: uncertain clinical or vital value, such as BP, pulse, respiratory rate, SpO2, or pain score.

Add a concise explanation in Notes whenever a cell is colored.

## Date And Time Sense

Use common sequence sense for Date and Time only. The goal is to reduce the user's review work while still showing assumptions clearly.

If a time appears impossible in row order, make the most reasonable date/time assumption and color the assumed cell blue. Example: if photo order shows `7AM -> 10AM -> 2PM -> 10PM -> 4PM`, and the final handwritten time could plausibly be `11PM`, enter `11PM`, color that Time cell blue, and note: `Assumed 11PM because original looked like 4PM but appears after 10PM in photo order.`

Also color blue when:

- time goes backward within the same date in source/photo order;
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

1. Identify all source photos/scans and convert unsupported formats such as HEIC to viewable copies if needed. Do not alter originals.
2. Read each photo in visible row order. Preserve photo order and source file for every row.
3. Build the `Review - Photo Order` sheet.
4. Apply the color rules: blue for date/time assumptions or sequence concerns, orange for uncertain clinical values.
5. Include a `Notes` sheet with the color legend, source list, and caution that handwritten clinical data needs human confirmation.
6. Ask the user to review colored cells. Do not create the final sorted sheet until the user confirms review is done.
7. After user confirmation, make a backup of the user-edited workbook before changing it.
8. Treat the user-edited workbook as the source of truth.
9. Re-check date/time sequence. If issues remain, highlight them blue and report them.
10. Create `Final - Sorted` sorted by Date then Time, or update the final workbook if the user explicitly wants a single sheet.
11. Verify Date cells display as human-readable dates such as `d/m`, not Excel serial numbers like `46175`.
12. Verify formulas/errors/rendering before final delivery.

## Review Checks

Before final delivery, check:

- no unsupported image files remain unread;
- all extracted rows have source photo and row number;
- time format is consistent;
- date display is human readable;
- blue/orange cells are cell-specific;
- Notes explains each assumption or uncertainty;
- final sorted sheet has no backward time jumps within the same date unless intentionally flagged.

## Safety

This skill produces a transcription aid, not a clinical authority. Tell the user that uncertain or clinically important readings should be verified against the original chart or by appropriate clinical staff.
