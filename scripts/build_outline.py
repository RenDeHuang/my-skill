"""Build dual-mode PPT outlines from extracted material summaries.

Modes:
- presentation: concise bullets + speaker notes
- self_explanatory: denser on-slide text with no speaker notes
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


def _dedupe_keep_order(items: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for item in items:
        key = item.strip()
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(key)
    return out


def _collect_points(summary: dict[str, Any], limit: int = 80) -> list[str]:
    raw_points: list[str] = []
    for doc in summary.get("documents", []):
        raw_points.extend(doc.get("key_points", []))

    cleaned: list[str] = []
    for point in raw_points:
        text = " ".join(point.split())
        if len(text) < 20:
            continue
        cleaned.append(text)

    points = _dedupe_keep_order(cleaned)
    return points[:limit]


def _truncate(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 1].rstrip() + "…"


def _agenda_items(mode: str) -> list[str]:
    if mode == "presentation":
        return [
            "研究背景与问题定义",
            "PCsRNAdb资源构建方法",
            "质量控制与方法学验证",
            "审稿意见与修订响应",
            "应用价值与下一步计划",
        ]
    return [
        "背景与需求：为什么需要统一的小RNA肿瘤资源库",
        "数据库构建：数据来源、处理流程与组织方式",
        "质量控制：关键校验策略与方法学一致性证据",
        "审稿响应：针对主要质疑的修订与补充实验",
        "价值与展望：临床与科研协同使用路径",
    ]


def _build_content_slide(
    title: str,
    points: list[str],
    mode: str,
    image_path: str | None,
) -> dict[str, Any]:
    if mode == "presentation":
        bullets = [_truncate(p, 72) for p in points[:4]]
        speaker_notes = "\n".join(f"- {p}" for p in points[:6])
    else:
        bullets = [_truncate(p, 180) for p in points[:5]]
        speaker_notes = ""

    slide = {
        "type": "content",
        "title": title,
        "bullets": bullets,
    }
    if speaker_notes:
        slide["speaker_notes"] = speaker_notes
    if image_path:
        slide["image"] = image_path
    return slide


def build_outline(summary: dict[str, Any], mode: str = "presentation") -> dict[str, Any]:
    if mode not in {"presentation", "self_explanatory"}:
        raise ValueError("mode must be 'presentation' or 'self_explanatory'")

    title = summary.get("project_title") or "研究材料汇报"
    subtitle = "Based on provided thesis, referee response, and publication draft"

    points = _collect_points(summary)
    if not points:
        points = ["未从材料中提取到足够文本，请检查输入文件。"]

    images = [img.get("path") for img in summary.get("images", []) if img.get("path")]

    def image_for(index: int) -> str | None:
        if not images:
            return None
        return images[index % len(images)]

    chunks: list[list[str]] = []
    chunk_size = 5 if mode == "self_explanatory" else 4
    for i in range(0, min(len(points), 24), chunk_size):
        chunk = points[i : i + chunk_size]
        if chunk:
            chunks.append(chunk)

    # Keep a minimum narrative depth even when source points are sparse.
    fallback_chunks = [
        [
            "研究问题聚焦在跨癌种小RNA资源整合的可用性与可比较性。",
            "目标是将分散数据转化为可检索、可解释、可复用的知识资产。",
        ],
        [
            "方法层面强调流程透明与质量控制，减少批次与平台差异带来的偏差。",
            "结果呈现遵循可追溯原则，便于研究者回查处理步骤与证据来源。",
        ],
        [
            "面对外部评审意见，重点补强方法对照、图表标注与结论边界描述。",
            "通过迭代修订提升数据库的可信度与社区可接受度。",
        ],
    ]
    while len(chunks) < 3:
        chunks.append(fallback_chunks[len(chunks)])

    section_titles = [
        "研究背景与目标",
        "数据库构建流程",
        "质量控制与验证",
        "审稿反馈与改进",
        "价值总结与下一步",
    ]

    slides: list[dict[str, Any]] = [
        {
            "type": "title",
            "title": title,
            "subtitle": subtitle,
        },
        {
            "type": "content",
            "title": "汇报结构",
            "bullets": _agenda_items(mode),
            "speaker_notes": "按五个模块推进，先讲问题定义，再讲方法、验证、反馈和落地。"
            if mode == "presentation"
            else "",
        },
    ]

    for idx, chunk in enumerate(chunks[:5]):
        slides.append(
            _build_content_slide(
                section_titles[idx],
                chunk,
                mode,
                image_for(idx),
            )
        )

    slides.append(
        {
            "type": "content",
            "title": "结论与行动项",
            "bullets": [
                "PCsRNAdb提供跨癌种小RNA资源整合能力，可支持机制探索与队列比较",
                "通过流程透明化与方法对照验证，增强结果可信度与可复现性",
                "下一步建议：扩展队列、增强临床注释、完善检索与可视化功能",
            ]
            if mode == "self_explanatory"
            else [
                "跨癌种小RNA资源：统一入口与可比较分析",
                "方法学可信度：流程与对照验证双重支撑",
                "下一步：扩队列、强注释、优交互",
            ],
            "speaker_notes": "结尾强调三件事：资源价值、可信度、后续路线。"
            if mode == "presentation"
            else "",
        }
    )

    return {
        "title": title,
        "subtitle": subtitle,
        "mode": mode,
        "theme": DEFAULT_THEME,
        "slides": slides,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Build dual-mode outline JSON")
    parser.add_argument("summary", help="Path to extracted summary JSON")
    parser.add_argument("output", help="Path to output outline JSON")
    parser.add_argument(
        "--mode",
        choices=["presentation", "self_explanatory"],
        default="presentation",
    )
    args = parser.parse_args()

    summary_path = Path(args.summary)
    output_path = Path(args.output)

    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    outline = build_outline(summary, mode=args.mode)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(outline, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Wrote outline: {output_path}")


if __name__ == "__main__":
    main()
