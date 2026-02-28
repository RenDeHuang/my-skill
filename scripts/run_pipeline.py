"""End-to-end pipeline: materials -> strategy-aligned outline(s) -> editable PPTX -> QA report."""

from __future__ import annotations

import argparse
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from build_outline import build_outline
from extract_materials import build_summary
from qa_deck import validate_deck
from render_editable_ppt import render_ppt_from_outline


def _slugify(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r"[^\w\u4e00-\u9fff-]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-")
    return text or "deck"


def _modes(requested: str) -> list[str]:
    if requested == "both":
        return ["presentation", "self_explanatory"]
    return [requested]


def _read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _default_strategy_path() -> Path:
    root = Path(__file__).resolve().parents[1]
    return root / "assets" / "strategy.template.json"


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate editable PPTX from source materials")
    parser.add_argument("--materials-dir", required=True, help="Folder with source docs/pdf/pptx")
    parser.add_argument(
        "--output-dir",
        default="/Users/huangrende/Desktop/ppt/outputs",
        help="Output folder for decks and intermediate JSON",
    )
    parser.add_argument(
        "--mode",
        choices=["presentation", "self_explanatory", "both"],
        default="both",
        help="Deck style mode",
    )
    parser.add_argument(
        "--deck-name",
        default="",
        help="Optional short name for output files (recommended)",
    )
    parser.add_argument("--template", default="", help="Optional template PPTX path")
    parser.add_argument("--theme", default="", help="Optional theme JSON path")
    parser.add_argument(
        "--strategy",
        default="",
        help="Optional strategy JSON path; defaults to assets/strategy.template.json when available",
    )
    args = parser.parse_args()

    materials_dir = Path(args.materials_dir).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    extracted_dir = output_dir / "extracted"
    summary = build_summary(materials_dir, extracted_dir)

    theme_override: dict[str, Any] = {}
    if args.theme:
        theme_override = _read_json(Path(args.theme).resolve())

    strategy_obj: dict[str, Any] | None = None
    if args.strategy:
        strategy_obj = _read_json(Path(args.strategy).resolve())
    else:
        default_strategy = _default_strategy_path()
        if default_strategy.exists():
            strategy_obj = _read_json(default_strategy)

    now = datetime.now().strftime("%Y%m%d-%H%M%S")
    slug_source = args.deck_name if args.deck_name else materials_dir.name
    slug = _slugify(slug_source)

    generated = []
    for mode in _modes(args.mode):
        outline = build_outline(summary, mode=mode, strategy=strategy_obj)
        if theme_override:
            merged = dict(outline.get("theme", {}))
            merged.update(theme_override)
            outline["theme"] = merged

        outline_path = output_dir / f"{slug}-{mode}-{now}.outline.json"
        outline_path.write_text(json.dumps(outline, ensure_ascii=False, indent=2), encoding="utf-8")

        deck_path = output_dir / f"{slug}-{mode}-{now}.pptx"
        template = Path(args.template).resolve() if args.template else None
        render_ppt_from_outline(outline, deck_path, template_path=template)

        qa_report = validate_deck(deck_path, outline)
        qa_path = output_dir / f"{slug}-{mode}-{now}.qa.json"
        qa_path.write_text(json.dumps(qa_report, ensure_ascii=False, indent=2), encoding="utf-8")

        generated.append(
            {
                "mode": mode,
                "outline": str(outline_path),
                "pptx": str(deck_path),
                "qa": str(qa_path),
                "qa_passed": qa_report.get("passed", False),
            }
        )
        print(f"Generated [{mode}] -> {deck_path}")
        print(f"QA [{mode}] -> {'PASS' if qa_report.get('passed') else 'FAIL'} ({qa_path})")

    report = {
        "materials_dir": str(materials_dir),
        "generated_at": datetime.now().isoformat(),
        "strategy_source": args.strategy or (str(_default_strategy_path()) if _default_strategy_path().exists() else ""),
        "generated": generated,
        "all_passed": all(item.get("qa_passed", False) for item in generated),
    }
    report_path = output_dir / f"generation-report-{now}.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote report: {report_path}")


if __name__ == "__main__":
    main()
