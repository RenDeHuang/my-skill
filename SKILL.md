---
name: editable-ppt-fusion
description: Use when generating or updating editable PPTX decks from mixed source materials (docx/pdf/pptx), especially when both speaker-led and self-explanatory modes are needed, with optional pinned template rendering, image insertion, and consistent visual style control.
---

# Editable PPT Fusion Skill

## Goal

Generate **fully editable** PowerPoint decks from source materials with a strict two-mode workflow:
- `presentation`: 面向演讲场景，页面聚焦 + 讲稿补充（`speaker_script`）
- `self_explanatory`: 面向自读场景，页面高信息密度 + 设计阐述（`design_rationale`）

This skill is aligned with three prompt specs in `/Users/huangrende/Desktop/ppt/自动PPT`:
- 先做策略化大纲与讲稿
- 自解释型页面知识卡片
- 按大纲落地可编辑成品并控制视觉规范

## Workflow Contract

1. **Extract facts** from source materials (论文、旧PPT、补充文档、图片)。
2. **Apply strategy** (`assets/strategy.template.json` or `--strategy`) to drive audience, goal, style, pacing.
3. **Build outline JSON** with markdown-aligned fields:
   - `slide_number`, `page_type`, `title`
   - `on_slide_content`, `visual_spec`
   - mode-specific: `speaker_script` or `design_rationale`
4. **Render editable PPTX** using text boxes and vector shapes (never full-slide raster output).
5. **Run QA checks** and emit `*.qa.json`:
   - no placeholders
   - every content slide has at least one `viz-*`
   - content slides include editable text frames

## Commands

### A) Dual-mode generation (recommended)

```bash
python /Users/huangrende/Desktop/ppt/editable-ppt-fusion/scripts/run_pipeline.py \
  --materials-dir /Users/huangrende/Desktop/ppt/毕业材料 \
  --output-dir /Users/huangrende/Desktop/ppt/outputs \
  --mode both \
  --deck-name 毕业材料-答辩版 \
  --strategy /Users/huangrende/Desktop/ppt/editable-ppt-fusion/assets/strategy.template.json \
  --theme /Users/huangrende/Desktop/ppt/editable-ppt-fusion/assets/theme-default.json
```

### B) Use pinned template

```bash
python /Users/huangrende/Desktop/ppt/editable-ppt-fusion/scripts/run_pipeline.py \
  --materials-dir /Users/huangrende/Desktop/ppt/毕业材料 \
  --output-dir /Users/huangrende/Desktop/ppt/outputs \
  --mode both \
  --deck-name 毕业材料-答辩版 \
  --template /absolute/path/to/template.pptx
```

## Output Files

Each run writes:
- `*.outline.json` (strategy-aligned outline per mode)
- `*.pptx` (editable deck per mode)
- `*.qa.json` (QA report per mode)
- `generation-report-*.json` (run summary, includes `qa_passed`)
- `extracted/summary.json` + extracted images

## Design Requirements Enforced

- 16:9 format for desktop reading.
- Cover/section/content page roles are explicit (`page_type`).
- No default placeholders allowed in final deck.
- Content pages include both text and simple visual elements (`viz-*`).
- Image policy is `material_only`: do not assume missing images.
- Citation policy defaults to `public_sources_only`.

## References

- Strategy input template: `assets/strategy.template.json`
- Outline contract: `assets/outline.schema.json`
- Mode mapping: `references/mode-mapping.md`
