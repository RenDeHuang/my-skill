import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from extract_materials import _is_rebuttal_file  # noqa: E402


class ContentPolicyTests(unittest.TestCase):
    def test_rebuttal_file_detection(self):
        self.assertTrue(_is_rebuttal_file("Response to referees.docx"))
        self.assertTrue(_is_rebuttal_file("审稿意见回复.docx"))
        self.assertTrue(_is_rebuttal_file("point-by-point rebuttal.pdf"))
        self.assertFalse(_is_rebuttal_file("中山大学硕士学位论文-黄仁德.docx"))
        self.assertFalse(_is_rebuttal_file("gkaf992.pdf"))


if __name__ == "__main__":
    unittest.main()
