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


if __name__ == "__main__":
    unittest.main()
