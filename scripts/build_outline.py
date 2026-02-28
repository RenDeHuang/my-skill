"""Build strategy-aligned dual-mode PPT outlines from extracted material summaries.

Modes:
- presentation: concise on-slide points + speaker scripts
- self_explanatory: denser on-slide text + design rationale
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

DEFAULT_THEME = {
    "font_name": "Noto Sans SC",
    "title_color": "1E3A8A",
    "body_color": "0F172A",
    "accent_color": "0EA5E9",
    "background_color": "F8FAFC",
}

DEFAULT_STRATEGY = {
    "speaker_role": "医学AI专家",
    "audience_profile": "一线医生",
    "core_goal": "知识传递与建立信任",
    "style_by_section": {
        "clinical": "专业严谨",
        "ai_principle": "生动科普",
    },
    "target_slide_count": 18,
    "max_minutes_per_slide": 1,
    "content_depth": "专业翔实",
    "require_chapter_dividers": True,
    "citation_policy": "public_sources_only",
}


def _dedupe_keep_order(items: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for item in items:
        key = " ".join(item.split())
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(key)
    return out


def _collect_points(summary: dict[str, Any], limit: int = 120) -> list[str]:
    points: list[str] = []
    for doc in summary.get("documents", []):
        for point in doc.get("key_points", []):
            text = " ".join(str(point).split())
            if len(text) < 15:
                continue
            points.append(text)
    return _dedupe_keep_order(points)[:limit]


def _truncate(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 1].rstrip() + "…"


def _normalize_strategy(raw: dict[str, Any] | None) -> dict[str, Any]:
    merged = dict(DEFAULT_STRATEGY)
    if raw:
        merged.update({k: v for k, v in raw.items() if v is not None})

    base_styles = dict(DEFAULT_STRATEGY["style_by_section"])
    extra_styles = merged.get("style_by_section") if isinstance(merged.get("style_by_section"), dict) else {}
    base_styles.update(extra_styles)
    merged["style_by_section"] = base_styles

    try:
        merged["target_slide_count"] = max(8, int(merged.get("target_slide_count", 18)))
    except (TypeError, ValueError):
        merged["target_slide_count"] = 18

    try:
        merged["max_minutes_per_slide"] = max(1, int(merged.get("max_minutes_per_slide", 1)))
    except (TypeError, ValueError):
        merged["max_minutes_per_slide"] = 1

    merged["require_chapter_dividers"] = bool(merged.get("require_chapter_dividers", True))
    merged["citation_policy"] = str(merged.get("citation_policy", "public_sources_only"))
    return merged


def _agenda_items(mode: str) -> list[str]:
    if mode == "presentation":
        return [
            "背景与痛点：为什么要做这件事",
            "资源构建方法：数据来源与处理流程",
            "质量控制与验证：可信度证据链",
            "审稿意见与修订：如何提升可复现性",
            "价值总结与下一步：落地路径",
        ]
    return [
        "背景与痛点：从临床和科研协同角度定义需求边界与核心问题",
        "资源构建方法：明确样本来源、处理流程、统计策略与输出组织方式",
        "质量控制与验证：通过一致性对照和可追溯记录建立结果可信性",
        "审稿意见与修订：围绕主要质疑补充证据并收敛结论表达边界",
        "价值总结与下一步：明确在科研复用和临床转化中的优先行动项",
    ]


def _section_chunks(points: list[str], mode: str) -> list[list[str]]:
    target_len = 3 if mode == "presentation" else 4
    chunks: list[list[str]] = []
    for idx in range(0, min(len(points), 28), target_len):
        chunk = points[idx : idx + target_len]
        if chunk:
            chunks.append(chunk)

    fallback = [
        [
            "研究问题聚焦在跨癌种小RNA资源的统一组织、可检索与可比较分析能力。",
            "目标是在保证证据可追溯的前提下，提升研究者获取关键结论的效率。",
            "需要兼顾临床可理解性与算法过程透明度，避免黑箱式结果展示。",
        ],
        [
            "构建流程覆盖数据标准化、质量筛选、特征计算与元信息对齐等关键步骤。",
            "在多队列、多平台场景中，必须通过一致性对照降低批次效应干扰。",
            "输出结构以可复用为前提，支持后续差异分析、功能注释与生存关联验证。",
        ],
        [
            "审稿修订聚焦方法学可解释性、图表标注规范和结果边界表达三类问题。",
            "通过补充对照实验和来源说明，增强结论的可靠性与外部可验证性。",
            "最终版本将复现路径写入文档，确保读者能重跑关键流程并复核结论。",
        ],
    ]
    while len(chunks) < 3:
        chunks.append(fallback[len(chunks)])
    return chunks


def _make_slide(
    *,
    mode: str,
    page_type: str,
    title: str,
    content: list[str] | None = None,
    visual_kind: str | None = None,
    section_style: str = "专业严谨",
    image_path: str | None = None,
    source_hint: str = "",
) -> dict[str, Any]:
    content = content or []
    dense = mode == "self_explanatory"

    on_slide_content = [
        _truncate(c, 190 if dense else 78)
        for c in content
    ]

    slide: dict[str, Any] = {
        "slide_number": 0,
        "page_type": page_type,
        "title": title,
        "section_style": section_style,
    }

    # Keep renderer compatibility while introducing markdown-aligned contract.
    if page_type == "cover":
        slide["type"] = "title"
        if on_slide_content:
            slide["subtitle"] = on_slide_content[0]
    elif page_type == "section_divider":
        slide["type"] = "section"
    else:
        slide["type"] = "content"
        slide["on_slide_content"] = on_slide_content
        slide["visual_spec"] = {
            "kind": visual_kind or "icon_list",
            "layout": "left_text_right_visual",
            "image_policy": "material_only",
            "image_path": image_path or "",
            "source": source_hint,
        }
        slide["bullets"] = on_slide_content

        if mode == "presentation":
            script = "；".join(content[:4])
            if not script:
                script = "按本页三点顺序讲解，先结论后证据，最后落到行动建议。"
            slide["speaker_script"] = script
            slide["speaker_notes"] = script
        else:
            rationale = (
                "本页采用结构化文本与简单信息图组合，目标是在无口头讲解前提下让读者"
                "独立理解关键结论、证据来源与逻辑关系。"
            )
            slide["design_rationale"] = rationale
            slide["speaker_notes"] = ""

    return slide


def build_outline(
    summary: dict[str, Any],
    mode: str = "presentation",
    strategy: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if mode not in {"presentation", "self_explanatory"}:
        raise ValueError("mode must be 'presentation' or 'self_explanatory'")

    strategy_obj = _normalize_strategy(strategy)
    title = summary.get("project_title") or "研究材料汇报"
    subtitle = (
        f"面向{strategy_obj['audience_profile']} | 角色: {strategy_obj['speaker_role']} | 目标: {strategy_obj['core_goal']}"
    )

    points = _collect_points(summary)
    chunks = _section_chunks(points, mode)
    images = [img.get("path") for img in summary.get("images", []) if img.get("path")]

    def image_for(index: int) -> str | None:
        if not images:
            return None
        return images[index % len(images)]

    slides: list[dict[str, Any]] = []

    slides.append(
        _make_slide(
            mode=mode,
            page_type="cover",
            title=title,
            content=[subtitle],
        )
    )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="agenda",
            title="汇报结构",
            content=_agenda_items(mode),
            visual_kind="three_column_compare",
            section_style="平衡",
            source_hint="结构规划",
        )
    )

    if strategy_obj.get("require_chapter_dividers", True):
        slides.append(
            _make_slide(
                mode=mode,
                page_type="section_divider",
                title="一、背景与问题定义",
                section_style=strategy_obj["style_by_section"].get("clinical", "专业严谨"),
            )
        )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="background",
            title="为什么需要统一的小RNA资源库",
            content=chunks[0],
            visual_kind="icon_list",
            section_style=strategy_obj["style_by_section"].get("clinical", "专业严谨"),
            image_path=image_for(0),
            source_hint="原始论文与项目材料",
        )
    )

    if strategy_obj.get("require_chapter_dividers", True):
        slides.append(
            _make_slide(
                mode=mode,
                page_type="section_divider",
                title="二、资源构建方法",
                section_style=strategy_obj["style_by_section"].get("ai_principle", "生动科普"),
            )
        )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="methodology",
            title="PCsRNAdb构建流程与关键步骤",
            content=chunks[1],
            visual_kind="two_step_flow",
            section_style=strategy_obj["style_by_section"].get("ai_principle", "生动科普"),
            image_path=image_for(1),
            source_hint="方法学章节",
        )
    )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="quality_control",
            title="质量控制与一致性验证",
            content=[
                "通过跨队列和跨工具对照验证关键指标，避免单一流程结论偏差。",
                "对样本覆盖率、注释完整性和统计稳定性设置明确阈值。",
                "所有关键图表均保留来源链路，便于复核与复现实验。",
            ],
            visual_kind="bar_compare",
            section_style="专业严谨",
            image_path=image_for(2),
            source_hint="结果与补充材料",
        )
    )

    if strategy_obj.get("require_chapter_dividers", True):
        slides.append(
            _make_slide(
                mode=mode,
                page_type="section_divider",
                title="三、审稿与修订",
                section_style="专业严谨",
            )
        )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="revision",
            title="审稿意见响应与修订价值",
            content=chunks[2],
            visual_kind="issue_solution_matrix",
            section_style="专业严谨",
            source_hint="rebuttal文稿",
        )
    )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="application",
            title="应用价值与落地场景",
            content=[
                "为机制研究提供统一入口，缩短从问题提出到候选靶点筛选的路径。",
                "支持与临床结局关联分析，帮助识别潜在预后生物标志物。",
                "为后续产品化接口与可视化平台建设提供结构化数据底座。",
            ],
            visual_kind="three_column_compare",
            section_style="平衡",
            source_hint="应用章节",
        )
    )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="conclusion",
            title="结论与下一步行动",
            content=[
                "结论一：跨癌种小RNA资源整合显著提升研究可比性与检索效率。",
                "结论二：透明流程与对照验证是建立信任的关键。",
                "行动项：扩展队列、增强临床注释、完善开放引用与复现实验入口。",
            ],
            visual_kind="timeline",
            section_style="专业严谨",
            source_hint="综合结论",
        )
    )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="qa",
            title="Q&A / Thank You",
            content=[
                "欢迎围绕方法、数据来源与应用场景继续讨论。",
                "如需复现流程，可按资料中的公开来源和版本说明执行。",
            ],
            visual_kind="icon_list",
            section_style="平衡",
            source_hint="结束页",
        )
    )

    for idx, slide in enumerate(slides, start=1):
        slide["slide_number"] = idx

    return {
        "title": title,
        "subtitle": subtitle,
        "mode": mode,
        "theme": DEFAULT_THEME,
        "strategy": strategy_obj,
        "slides": slides,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Build strategy-aligned dual-mode outline JSON")
    parser.add_argument("summary", help="Path to extracted summary JSON")
    parser.add_argument("output", help="Path to output outline JSON")
    parser.add_argument(
        "--mode",
        choices=["presentation", "self_explanatory"],
        default="presentation",
    )
    parser.add_argument("--strategy", default="", help="Optional strategy JSON path")
    args = parser.parse_args()

    summary_path = Path(args.summary)
    output_path = Path(args.output)

    strategy_obj = None
    if args.strategy:
        strategy_obj = json.loads(Path(args.strategy).read_text(encoding="utf-8"))

    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    outline = build_outline(summary, mode=args.mode, strategy=strategy_obj)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(outline, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Wrote outline: {output_path}")


if __name__ == "__main__":
    main()
