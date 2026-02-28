"""Extract key text and reusable images from research material folders."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from typing import Any

from docx import Document
from PIL import Image
from pypdf import PdfReader

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}
SKIP_PATTERNS = [
    "原创性声明",
    "学位论文使用授权声明",
    "学位论文作者签名",
    "答辩委员会",
]
REBUTTAL_FILE_PATTERNS = [
    "response to referees",
    "response to reviewers",
    "rebuttal",
    "point-by-point",
    "审稿",
    "回复",
    "referee",
    "reviewer",
]


def _clean_text(text: str) -> str:
    return " ".join((text or "").replace("\u3000", " ").split())


def _looks_useful(line: str) -> bool:
    if len(line) < 18:
        return False
    return not any(pat in line for pat in SKIP_PATTERNS)


def _is_rebuttal_file(name: str) -> bool:
    lowered = str(name).lower()
    return any(pat in lowered for pat in REBUTTAL_FILE_PATTERNS)


def _extract_docx_points(path: Path, max_points: int = 120) -> list[str]:
    doc = Document(str(path))
    points = []
    for p in doc.paragraphs:
        line = _clean_text(p.text)
        if _looks_useful(line):
            points.append(line)
        if len(points) >= max_points:
            break
    return points


def _extract_pdf_points(path: Path, max_points: int = 80) -> list[str]:
    reader = PdfReader(str(path))
    points: list[str] = []

    for page in reader.pages:
        text = _clean_text(page.extract_text() or "")
        if not text:
            continue
        # Split long text blocks into sentence-like units.
        for chunk in re.split(r"(?<=[。.!?])\s+", text):
            chunk = _clean_text(chunk)
            if _looks_useful(chunk):
                points.append(chunk)
            if len(points) >= max_points:
                return points

    return points


def _extract_zip_images(source: Path, output_dir: Path) -> list[Path]:
    if source.suffix.lower() not in {".pptx", ".docx"}:
        return []

    dest_root = output_dir / source.stem
    dest_root.mkdir(parents=True, exist_ok=True)

    prefixes = ["ppt/media/", "word/media/"]
    extracted: list[Path] = []

    with zipfile.ZipFile(source, "r") as zf:
        for name in zf.namelist():
            if not any(name.startswith(prefix) for prefix in prefixes):
                continue
            if name.endswith("/"):
                continue
            ext = Path(name).suffix.lower()
            if ext not in IMAGE_EXTS:
                continue

            data = zf.read(name)
            candidate = dest_root / Path(name).name
            if candidate.exists():
                candidate = dest_root / f"{Path(name).stem}_{len(extracted)+1}{ext}"
            candidate.write_bytes(data)
            extracted.append(candidate)

    return extracted


def _image_meta(path: Path) -> dict[str, Any] | None:
    try:
        with Image.open(path) as img:
            width, height = img.size
        if width < 320 or height < 180:
            return None
        return {"path": str(path), "width": width, "height": height}
    except Exception:
        return None


def _derive_title(documents: list[dict[str, Any]]) -> str:
    candidates: list[tuple[float, str]] = []
    for doc in documents:
        kind = doc.get("kind")
        for line in doc.get("key_points", []):
            if "PCsRNAdb" not in line:
                continue
            cleaned = _clean_text(line)
            if len(cleaned) > 120:
                continue

            score = 0.0
            if kind == "docx":
                score += 2.0
            if "综合资源库" in cleaned or "comprehensive resource" in cleaned.lower():
                score += 3.0
            if "一个涵盖" in cleaned or "across cancers" in cleaned.lower():
                score += 2.0
            score -= len(cleaned) / 200.0
            candidates.append((score, cleaned[:90]))

    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]
    return "毕业论文与论文工作汇报"


def build_summary(
    materials_dir: Path,
    output_dir: Path,
    exclude_rebuttal: bool = True,
) -> dict[str, Any]:
    materials_dir = materials_dir.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    images_dir = output_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    documents: list[dict[str, Any]] = []
    image_paths: list[Path] = []

    files = sorted([p for p in materials_dir.iterdir() if p.is_file()])
    for path in files:
        suffix = path.suffix.lower()
        if exclude_rebuttal and _is_rebuttal_file(path.name):
            continue

        if suffix == ".docx":
            points = _extract_docx_points(path)
            documents.append(
                {
                    "file": path.name,
                    "kind": "docx",
                    "key_points": points,
                    "point_count": len(points),
                }
            )
            image_paths.extend(_extract_zip_images(path, images_dir))
        elif suffix == ".pdf":
            points = _extract_pdf_points(path)
            documents.append(
                {
                    "file": path.name,
                    "kind": "pdf",
                    "key_points": points,
                    "point_count": len(points),
                }
            )
        elif suffix == ".pptx":
            documents.append(
                {
                    "file": path.name,
                    "kind": "pptx",
                    "key_points": [
                        "该文件作为既有演示材料输入，可用于抽取视觉风格与图像素材。"
                    ],
                    "point_count": 1,
                }
            )
            image_paths.extend(_extract_zip_images(path, images_dir))

    # Prefer high-resolution images.
    image_meta = [_image_meta(p) for p in image_paths]
    images = [m for m in image_meta if m]
    images.sort(key=lambda x: x["width"] * x["height"], reverse=True)

    summary = {
        "materials_dir": str(materials_dir),
        "project_title": _derive_title(documents),
        "documents": documents,
        "images": images[:30],
    }

    summary_path = output_dir / "summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract material summary for PPT generation")
    parser.add_argument("materials_dir", help="Input folder with source files")
    parser.add_argument("output_dir", help="Directory for extracted summary/images")
    parser.add_argument(
        "--include-rebuttal",
        action="store_true",
        help="Include rebuttal/reviewer-response files (default excludes them)",
    )
    args = parser.parse_args()

    summary = build_summary(
        Path(args.materials_dir),
        Path(args.output_dir),
        exclude_rebuttal=not args.include_rebuttal,
    )
    print(f"Summary created with {len(summary['documents'])} documents and {len(summary['images'])} images")


if __name__ == "__main__":
    main()
