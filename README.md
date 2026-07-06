# Codex Skills

Personal Codex skills collected in one repository.

## Structure

Each skill lives under `skills/<skill-name>/` and should contain:

- `SKILL.md`: required skill instructions and trigger metadata.
- `agents/openai.yaml`: optional UI metadata.
- `scripts/`: optional deterministic helper scripts.
- `references/`: optional supporting docs loaded only when needed.
- `assets/`: optional templates or reusable files.

## Skills

- `vitals-chart-extractor`: Extract handwritten or image-based clinical vitals observation charts into review-first Excel workbooks.

## Install Locally

Copy a skill folder into your Codex skills directory:

```bash
mkdir -p ~/.codex/skills
cp -R skills/vitals-chart-extractor ~/.codex/skills/
```

Then start a new Codex session and invoke it with:

```text
Use $vitals-chart-extractor on these clinical vitals chart images and make me an Excel file.
```

## Development Notes

Keep skill folders lean. Put reusable code in `scripts/` only when it improves reliability or reduces repeated manual work. For clinical handwriting tasks, scripts should handle workbook mechanics, validation, formatting, sorting, backups, and graph-selection sheet setup; Codex should still make the visual judgement calls and clearly mark assumptions for human review.
