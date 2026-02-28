---
name: editable-ppt-fusion
description: Use when generating or updating editable PPTX decks from mixed source materials (docx/pdf/pptx), especially when both speaker-led and self-explanatory modes are needed, with optional pinned template rendering, image insertion, and consistent visual style control.
---

# Editable PPT Fusion Skill

## Overview

Use this skill to produce **editable** PowerPoint decks from source materials.
It supports two modes in one workflow:
- `presentation`: concise slides with speaker notes
- `self_explanatory`: denser on-slide text that can be read without narration

## Workflow

1. Extract source facts and candidate images from material files.
2. Generate dual-mode outline JSON files.
3. Render editable PPTX using text boxes, shapes, and image/chart panels (not screenshot export).
4. Optionally apply a pinned template via `--template`.
5. Optionally override style via `--theme`.

## Commands

Run end-to-end generation:

```bash
python /Users/huangrende/Desktop/ppt/editable-ppt-fusion/scripts/run_pipeline.py \
  --materials-dir /Users/huangrende/Desktop/ppt/毕业材料 \
  --output-dir /Users/huangrende/Desktop/ppt/outputs \
  --mode both \
  --deck-name 毕业材料 \
  --theme /Users/huangrende/Desktop/ppt/editable-ppt-fusion/assets/theme-default.json
```

Use a pinned template:

```bash
python /Users/huangrende/Desktop/ppt/editable-ppt-fusion/scripts/run_pipeline.py \
  --materials-dir /Users/huangrende/Desktop/ppt/毕业材料 \
  --output-dir /Users/huangrende/Desktop/ppt/outputs \
  --mode both \
  --deck-name 毕业材料 \
  --template /absolute/path/to/template.pptx
```

## Outputs

Generated files are saved to the output directory:
- `*.outline.json` (one per mode)
- `*.pptx` (editable deck, one per mode)
- `generation-report-*.json` (run manifest)
- `extracted/summary.json` and extracted images

## Design Rules

- Keep PPT editable: do not rasterize whole slides into images for PPTX export.
- Remove all template placeholders (`单击此处添加标题/文本`) before writing slide content.
- Every content slide must include at least one `viz-*` visual component (panel, chart-like bars, or image card).
- Prefer material-derived images; skip low-resolution assets.
- Keep typography and colors consistent through the theme file.
- Use `presentation` mode when narration is available.
- Use `self_explanatory` mode for async reading and archival decks.

## QA Checks (Required)

After rendering, verify:
- No placeholder shapes remain in any slide.
- Every content slide has at least one shape named `viz-*`.
- Slide text is editable in PowerPoint (text boxes, not flattened images).

## References

- Mode mapping and differences: `references/mode-mapping.md`
- Theme fields and defaults: `assets/theme-default.json`
