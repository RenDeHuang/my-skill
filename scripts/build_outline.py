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

REBUTTAL_PATTERNS = [
    "response to referees",
    "response to reviewers",
    "rebuttal",
    "reviewer",
    "referee",
    "审稿",
    "回复",
]

NOISE_PATTERNS = [
    "email:",
    "correspondence",
    "to whom correspondence",
    "https://",
    "http://",
    "doi.org",
    "advance access publication",
    "the first four authors",
    "department of",
    "school of",
    "author@example.com",
    "本人完全了解中山大学有关保留",
    "学位论文使用授权声明",
    "原创性声明",
    "答辩委员会",
    "作者签名",
]

TOPIC_HINT_KEYWORDS = [
    "sncrna",
    "mirna",
    "pirna",
    "tdr",
    "rrf",
    "cancer",
    "pan-cancer",
    "database",
    "pcsrnadb",
    "pipeline",
    "method",
    "analysis",
    "result",
    "survival",
    "differential",
    "expression",
    "tumor",
]

BACKGROUND_KEYWORDS = [
    "背景",
    "意义",
    "目的",
    "痛点",
    "现状",
    "introduction",
    "background",
    "challenge",
]
METHOD_KEYWORDS = [
    "方法",
    "流程",
    "数据",
    "数据库",
    "构建",
    "pipeline",
    "method",
    "dataset",
    "processing",
    "analysis workflow",
]
RESULT_KEYWORDS = [
    "结果",
    "发现",
    "提升",
    "显著",
    "验证",
    "result",
    "finding",
    "survival",
    "differential",
    "performance",
    "association",
]
OUTLOOK_KEYWORDS = [
    "展望",
    "未来",
    "局限",
    "下一步",
    "discussion",
    "conclusion",
    "future",
    "limitation",
]

SECTION_FALLBACKS = {
    "background": [
        "本研究关注跨癌种小RNA资源分散、口径不一致导致的分析复用困难。",
        "核心目标是建立统一、可追溯、可比较的数据资源体系，降低重复整理成本。",
        "研究问题围绕临床解释需求与算法处理规范之间的协同展开。",
    ],
    "methodology": [
        "方法上采用标准化流程完成数据收集、质量筛选、注释映射和指标计算。",
        "通过统一元数据结构连接样本、癌种、分子类型与分析结果，保证检索一致性。",
        "关键步骤保留版本信息和处理参数，支撑复现与后续迭代。",
    ],
    "results": [
        "结果显示资源库能够稳定支持跨队列比较，并提供可解释的差异表达线索。",
        "在代表性任务中，关键指标表现与既有研究趋势一致，验证了流程可靠性。",
        "平台化组织显著提升了从问题提出到证据定位的效率。",
    ],
    "outlook": [
        "下一步将扩展更多公开队列并持续完善临床注释字段。",
        "计划增强交互可视化与结果导出能力，提升一线研究场景可用性。",
        "后续将围绕转化价值开展更系统的外部验证与协作。",
    ],
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


def _truncate(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 1].rstrip() + "…"


def _is_rebuttal_source(name: str) -> bool:
    lowered = str(name).lower()
    return any(pattern in lowered for pattern in REBUTTAL_PATTERNS)


def _looks_rebuttal_text(text: str) -> bool:
    lowered = str(text).lower()
    return any(pattern in lowered for pattern in REBUTTAL_PATTERNS)


def _contains_cjk(text: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in text)


def _looks_noise_text(text: str) -> bool:
    lowered = str(text).lower()
    if any(pattern in lowered for pattern in NOISE_PATTERNS):
        return True
    if lowered.count("@") > 0:
        return True
    if len(text) > 420:
        return True
    return False


def _looks_contentful_text(text: str) -> bool:
    if _looks_noise_text(text):
        return False
    lowered = text.lower()
    if _contains_cjk(text):
        return True
    return any(keyword in lowered for keyword in TOPIC_HINT_KEYWORDS)


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


def _collect_points(summary: dict[str, Any], limit: int = 180) -> list[str]:
    points: list[str] = []
    documents = list(summary.get("documents", []))

    def _doc_priority(doc: dict[str, Any]) -> tuple[int, int]:
        filename = str(doc.get("file", ""))
        kind = str(doc.get("kind", "")).lower()
        if "学位论文" in filename or "毕业论文" in filename:
            return (0, 0)
        if kind == "docx":
            return (1, 0)
        if kind == "pdf":
            return (2, 0)
        return (3, 0)

    for doc in sorted(documents, key=_doc_priority):
        filename = str(doc.get("file", ""))
        kind = str(doc.get("kind", "")).lower()

        if _is_rebuttal_source(filename):
            continue
        if kind == "pptx":
            # Existing PPT is used for style/image references, not textual narrative.
            continue

        for point in doc.get("key_points", []):
            text = " ".join(str(point).split())
            if len(text) < 15:
                continue
            if "既有演示材料输入" in text:
                continue
            if _looks_rebuttal_text(text):
                continue
            if not _looks_contentful_text(text):
                continue
            points.append(text)

    return _dedupe_keep_order(points)[:limit]


def _agenda_items(mode: str) -> list[str]:
    if mode == "presentation":
        return [
            "研究背景：问题定义与研究动机",
            "研究方法：技术路线与实现流程",
            "研究结果：核心发现与验证证据",
            "研究展望：价值总结与后续计划",
        ]
    return [
        "研究背景：说明临床与科研场景中数据分散、口径不统一带来的分析障碍与研究动机",
        "研究方法：交代样本来源、处理流程、统计策略和数据组织方式，确保可复现与可追溯",
        "研究结果：展示关键发现、验证思路与代表性分析结论，形成清晰证据链",
        "研究展望：总结研究价值、已知局限与下一阶段扩展方向，明确落地路径",
    ]


def _classify_point(point: str) -> str | None:
    lowered = point.lower()
    if any(keyword in lowered for keyword in BACKGROUND_KEYWORDS):
        return "background"
    if any(keyword in lowered for keyword in METHOD_KEYWORDS):
        return "methodology"
    if any(keyword in lowered for keyword in RESULT_KEYWORDS):
        return "results"
    if any(keyword in lowered for keyword in OUTLOOK_KEYWORDS):
        return "outlook"
    return None


def _section_payloads(points: list[str], mode: str) -> dict[str, list[str]]:
    size = 3 if mode == "presentation" else 4
    buckets = {
        "background": [],
        "methodology": [],
        "results": [],
        "outlook": [],
    }

    for point in points:
        section = _classify_point(point)
        if section:
            buckets[section].append(point)
            continue

        # Unlabeled points prefer results/methods for report effectiveness.
        if len(buckets["results"]) <= len(buckets["methodology"]):
            buckets["results"].append(point)
        else:
            buckets["methodology"].append(point)

    def _point_quality(point: str) -> int:
        score = 0
        lowered = point.lower()
        if _contains_cjk(point):
            score += 3
        if 30 <= len(point) <= 190:
            score += 1
        if any(keyword in lowered for keyword in TOPIC_HINT_KEYWORDS):
            score += 1
        if any(noise in lowered for noise in ["copyright", "all rights reserved"]):
            score -= 2
        return score

    for key in buckets:
        buckets[key] = sorted(buckets[key], key=_point_quality, reverse=True)

    payloads: dict[str, list[str]] = {}
    for section, fallback in SECTION_FALLBACKS.items():
        chosen = buckets[section][:size]
        fallback_idx = 0
        while len(chosen) < size and fallback_idx < len(fallback):
            chosen.append(fallback[fallback_idx])
            fallback_idx += 1
        payloads[section] = chosen

    return payloads


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

    on_slide_content = [_truncate(c, 190 if dense else 78) for c in content]

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
    title = summary.get("project_title") or "毕业论文与论文工作汇报"
    subtitle = (
        f"面向{strategy_obj['audience_profile']} | 角色: {strategy_obj['speaker_role']} | 目标: {strategy_obj['core_goal']}"
    )

    points = _collect_points(summary)
    sections = _section_payloads(points, mode)
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
                title="一、研究背景",
                section_style=strategy_obj["style_by_section"].get("clinical", "专业严谨"),
            )
        )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="background",
            title="研究背景与问题定义",
            content=sections["background"],
            visual_kind="icon_list",
            section_style=strategy_obj["style_by_section"].get("clinical", "专业严谨"),
            image_path=image_for(0),
            source_hint="论文背景章节",
        )
    )

    if strategy_obj.get("require_chapter_dividers", True):
        slides.append(
            _make_slide(
                mode=mode,
                page_type="section_divider",
                title="二、研究方法",
                section_style=strategy_obj["style_by_section"].get("ai_principle", "生动科普"),
            )
        )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="methodology",
            title="研究方法与技术路线",
            content=sections["methodology"],
            visual_kind="two_step_flow",
            section_style=strategy_obj["style_by_section"].get("ai_principle", "生动科普"),
            image_path=image_for(1),
            source_hint="论文方法章节",
        )
    )

    if strategy_obj.get("require_chapter_dividers", True):
        slides.append(
            _make_slide(
                mode=mode,
                page_type="section_divider",
                title="三、研究结果",
                section_style="专业严谨",
            )
        )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="results",
            title="研究结果与核心发现",
            content=sections["results"],
            visual_kind="bar_compare",
            section_style="专业严谨",
            image_path=image_for(2),
            source_hint="论文结果章节",
        )
    )

    if strategy_obj.get("require_chapter_dividers", True):
        slides.append(
            _make_slide(
                mode=mode,
                page_type="section_divider",
                title="四、研究展望",
                section_style="平衡",
            )
        )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="outlook",
            title="研究展望与后续工作",
            content=sections["outlook"],
            visual_kind="timeline",
            section_style="平衡",
            source_hint="总结与展望",
        )
    )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="conclusion",
            title="结论与答辩要点",
            content=[
                "本研究完成了跨癌种小RNA资源整合的框架构建，并形成可复用数据底座。",
                "方法与结果围绕可追溯、可比较、可解释三个维度建立证据链。",
                "答辩建议聚焦研究价值、方法可信性与未来扩展三条主线。",
            ],
            visual_kind="three_column_compare",
            section_style="专业严谨",
            source_hint="答辩收束页",
        )
    )

    slides.append(
        _make_slide(
            mode=mode,
            page_type="qa",
            title="Q&A / Thank You",
            content=[
                "欢迎围绕研究设计、统计策略与临床价值提出问题。",
                "如需复现流程，可按论文与文章中的公开来源和版本说明执行。",
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
