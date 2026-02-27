"""Render editable PPTX from a structured outline JSON."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


def _hex_to_rgb(color: str, fallback: str) -> RGBColor:
    value = (color or fallback).strip().lstrip("#")
    if len(value) != 6:
        value = fallback
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def _apply_background(slide, background_color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _hex_to_rgb(background_color, "FFFFFF")


def _find_layout(pres: Presentation, preferred_index: int, fallback_index: int = 6):
    if preferred_index < len(pres.slide_layouts):
        return pres.slide_layouts[preferred_index]
    if fallback_index < len(pres.slide_layouts):
        return pres.slide_layouts[fallback_index]
    return pres.slide_layouts[0]


def _clear_all_existing_slides(pres: Presentation) -> None:
    # python-pptx does not expose a public remove-all API; this pattern is
    # the stable internal approach used for template-based regeneration.
    slide_id_list = list(pres.slides._sldIdLst)  # pylint: disable=protected-access
    for slide_id in slide_id_list:
        rel_id = slide_id.rId
        pres.part.drop_rel(rel_id)
        pres.slides._sldIdLst.remove(slide_id)  # pylint: disable=protected-access


def _set_run_style(run, font_name: str, size: int, color: str, bold: bool = False) -> None:
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = _hex_to_rgb(color, "0F172A")


def _add_title_slide(pres: Presentation, spec: dict[str, Any], theme: dict[str, Any]) -> None:
    slide = pres.slides.add_slide(_find_layout(pres, 0, 0))
    _apply_background(slide, theme.get("background_color", "F8FAFC"))

    title_text = spec.get("title", "Untitled")
    subtitle_text = spec.get("subtitle", "")

    if slide.shapes.title is not None:
        title_shape = slide.shapes.title
        title_shape.text = title_text
        p = title_shape.text_frame.paragraphs[0]
        if p.runs:
            _set_run_style(
                p.runs[0],
                theme.get("font_name", "Calibri"),
                44,
                theme.get("title_color", "1E3A8A"),
                bold=True,
            )

    if len(slide.placeholders) > 1:
        subtitle = slide.placeholders[1]
        subtitle.text = subtitle_text
        p = subtitle.text_frame.paragraphs[0]
        if p.runs:
            _set_run_style(
                p.runs[0],
                theme.get("font_name", "Calibri"),
                20,
                theme.get("body_color", "0F172A"),
            )


def _add_content_slide(pres: Presentation, spec: dict[str, Any], theme: dict[str, Any]) -> None:
    slide = pres.slides.add_slide(_find_layout(pres, 6, 1))
    _apply_background(slide, theme.get("background_color", "F8FAFC"))

    # Title box
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.2), Inches(0.8))
    tf_title = title_box.text_frame
    tf_title.clear()
    p_title = tf_title.paragraphs[0]
    run_title = p_title.add_run()
    run_title.text = spec.get("title", "")
    _set_run_style(
        run_title,
        theme.get("font_name", "Calibri"),
        30,
        theme.get("title_color", "1E3A8A"),
        bold=True,
    )

    image_path = spec.get("image")
    has_image = isinstance(image_path, str) and Path(image_path).exists()

    text_x = Inches(0.8)
    text_y = Inches(1.3)
    text_w = Inches(7.5 if has_image else 11.6)
    text_h = Inches(5.8)

    body_box = slide.shapes.add_textbox(text_x, text_y, text_w, text_h)
    tf_body = body_box.text_frame
    tf_body.clear()
    tf_body.word_wrap = True

    bullets = spec.get("bullets", [])
    for idx, bullet in enumerate(bullets):
        p = tf_body.paragraphs[0] if idx == 0 else tf_body.add_paragraph()
        p.text = bullet
        p.level = 0
        p.space_after = Pt(10)
        if p.runs:
            _set_run_style(
                p.runs[0],
                theme.get("font_name", "Calibri"),
                20 if spec.get("mode") == "presentation" else 18,
                theme.get("body_color", "0F172A"),
            )

    accent = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        Inches(0.6),
        Inches(1.18),
        Inches(12.2),
        Inches(0.05),
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = _hex_to_rgb(theme.get("accent_color", "0EA5E9"), "0EA5E9")
    accent.line.fill.background()

    if has_image:
        slide.shapes.add_picture(str(Path(image_path)), Inches(8.5), Inches(1.45), Inches(4.2), Inches(4.9))

    speaker_notes = spec.get("speaker_notes")
    if speaker_notes:
        notes = slide.notes_slide.notes_text_frame
        notes.clear()
        notes.text = speaker_notes


def _add_section_slide(pres: Presentation, spec: dict[str, Any], theme: dict[str, Any]) -> None:
    slide = pres.slides.add_slide(_find_layout(pres, 6, 0))
    _apply_background(slide, theme.get("title_color", "1E3A8A"))

    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(11.8), Inches(2.0))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = spec.get("title", "章节")
    _set_run_style(run, theme.get("font_name", "Calibri"), 44, "FFFFFF", bold=True)


def render_ppt_from_outline(
    outline: dict[str, Any],
    output_path: Path,
    template_path: Path | None = None,
) -> None:
    if template_path is not None and template_path.exists():
        pres = Presentation(str(template_path))
        _clear_all_existing_slides(pres)
    else:
        pres = Presentation()

    # enforce 16:9
    pres.slide_width = Inches(13.333)
    pres.slide_height = Inches(7.5)

    theme = outline.get("theme", {})
    slides = outline.get("slides", [])

    for idx, slide_spec in enumerate(slides):
        slide_type = slide_spec.get("type", "content")
        enriched = dict(slide_spec)
        enriched.setdefault("mode", outline.get("mode", "presentation"))

        if idx == 0 and slide_type == "title":
            _add_title_slide(pres, enriched, theme)
        elif slide_type == "section":
            _add_section_slide(pres, enriched, theme)
        else:
            _add_content_slide(pres, enriched, theme)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    pres.save(str(output_path))


def main() -> None:
    parser = argparse.ArgumentParser(description="Render editable PPTX from outline JSON")
    parser.add_argument("outline", help="Path to outline JSON")
    parser.add_argument("output", help="Path to output PPTX")
    parser.add_argument("--template", help="Optional template PPTX", default="")
    args = parser.parse_args()

    outline = json.loads(Path(args.outline).read_text(encoding="utf-8"))
    template = Path(args.template) if args.template else None
    render_ppt_from_outline(outline, Path(args.output), template_path=template)
    print(f"Wrote editable PPTX: {args.output}")


if __name__ == "__main__":
    main()
