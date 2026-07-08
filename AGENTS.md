# Repository Instructions

This repository contains personal Codex skills. Keep skill folders lean and focused on instructions or reusable helpers that directly improve the skill.

## Current Skills

- Internal skill name: `vitals-chart-extractor`
  - User-facing display name: `Clinical Vitals Chart Extractor`
  - Purpose: extract handwritten or image-based clinical vitals observation charts into review-first Excel workbooks, then create final sorted and graph-selection sheets.
- Internal skill name: `vital-graph-builder`
  - User-facing display name: `Clinical Vital Graph Builder`
  - Purpose: read the ticked `Graph?` rows on the vital graph-selection sheets and draw native Excel charts on each vital's sheet - a BP double-arrow chart (error-bar based) and single-colour line charts, all with data tables and visible axes.

## Wording

Use neutral clinical-image language. Prefer:

- `clinical vitals chart images`
- `clinical observation chart images`
- `provided chart captures`
- `source image`
- `source file`
- `source order`

Avoid wording that sounds like encouraging photography of charts, such as:

- `photographed hospital vital-sign charts`
- `chart photos`
- `photo order`
- `Source Photo`

## Workbook Workflow

The vitals skill should produce:

1. `Review - Source Order`: source-order sheet for human checking.
2. `Final - Sorted`: confirmed sheet sorted by Date then Time.
3. Vital graph-selection sheets: `BP`, `Heart Rate`, `Respiratory`, `SpO2`, and `Pain Score`.
4. `Notes`: short practical legend/instructions only.

Graph-selection sheets should use real clickable Excel checkbox cells where available. The `Graph?` value must be readable later as `TRUE` or `FALSE`.

Keep normal Excel gridlines visible unless the user explicitly requests a print-style workbook.
