import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from build_outline import build_outline  # noqa: E402


class OutlineTests(unittest.TestCase):
    @staticmethod
    def _sample_summary():
        return {
            "project_title": "PCsRNAdb materials",
            "documents": [
                {
                    "file": "thesis.docx",
                    "kind": "docx",
                    "key_points": [
                        "PCsRNAdb is a comprehensive resource of small noncoding RNAs across cancers.",
                        "The dataset covers pan-cancer cohorts and multiple RNA categories.",
                        "Quality control and transparent metadata are important design principles.",
                    ],
                },
                {
                    "file": "review_response.docx",
                    "kind": "docx",
                    "key_points": [
                        "Reviewers asked about pipeline validation and consistency across projects.",
                        "Authors compared the pipeline with miRDeep2 and reported strong correlation.",
                    ],
                },
            ],
            "images": [
                {"path": "/tmp/image1.png", "width": 1200, "height": 800},
                {"path": "/tmp/image2.png", "width": 1000, "height": 700},
            ],
        }

    def test_presentation_outline_contains_speaker_notes(self):
        outline = build_outline(self._sample_summary(), mode="presentation")

        self.assertEqual(outline["mode"], "presentation")
        self.assertGreaterEqual(len(outline["slides"]), 6)
        self.assertTrue(
            any(s.get("speaker_notes") for s in outline["slides"] if s["type"] == "content")
        )

    def test_self_explanatory_outline_has_dense_text_without_notes(self):
        outline = build_outline(self._sample_summary(), mode="self_explanatory")

        self.assertEqual(outline["mode"], "self_explanatory")
        content_slides = [s for s in outline["slides"] if s["type"] == "content"]
        self.assertTrue(content_slides)
        self.assertTrue(all(not s.get("speaker_notes") for s in content_slides))
        self.assertTrue(any(len(" ".join(s.get("bullets", []))) > 140 for s in content_slides))

    def test_outline_uses_thesis_arc_and_excludes_rebuttal_narrative(self):
        summary = {
            "project_title": "毕业论文汇报",
            "documents": [
                {
                    "file": "中山大学硕士学位论文-黄仁德.docx",
                    "kind": "docx",
                    "key_points": [
                        "研究背景：当前缺乏跨癌种小RNA统一资源，影响可比性分析。",
                        "研究方法：构建标准化处理流程并建立统一注释体系。",
                        "研究结果：在多癌种队列中识别到稳定差异表达信号并完成生存关联分析。",
                        "研究展望：扩展样本规模并增强临床转化应用。",
                        "Email: author@example.com Correspondence may also be addressed to someone.",
                    ],
                },
                {
                    "file": "Response to referees.docx",
                    "kind": "docx",
                    "key_points": [
                        "Reviewer 1 requested additional validation experiments.",
                        "Response: we clarified comments point by point.",
                    ],
                },
            ],
            "images": [],
        }

        outline = build_outline(summary, mode="presentation")
        titles = " ".join(s.get("title", "") for s in outline["slides"])
        content = " ".join(" ".join(s.get("on_slide_content", [])) for s in outline["slides"])
        merged = f"{titles} {content}".lower()

        self.assertIn("研究背景", titles)
        self.assertIn("研究方法", titles)
        self.assertIn("研究结果", titles)
        self.assertIn("研究展望", titles)
        self.assertNotIn("审稿", merged)
        self.assertNotIn("reviewer", merged)
        self.assertNotIn("email:", merged)


if __name__ == "__main__":
    unittest.main()
