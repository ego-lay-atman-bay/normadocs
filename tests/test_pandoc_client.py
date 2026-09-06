"""
Tests for Pandoc Client.
"""

import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from docx import Document
from docx.enum.dml import MSO_COLOR_TYPE

from normadocs.pandoc_client import REFERENCE_DOCX, PandocRunner


class TestPandocClient(unittest.TestCase):
    @patch("subprocess.run")
    def test_run_success(self, mock_run):
        mock_result = MagicMock()
        mock_result.returncode = 0
        mock_run.return_value = mock_result

        runner = PandocRunner()
        success = runner.run("# Title", "output.docx")

        self.assertTrue(success)
        mock_run.assert_called_once()

        args, _ = mock_run.call_args
        cmd = args[0]
        self.assertTrue(cmd[0].endswith("pandoc") or cmd[0] == "pandoc")
        self.assertIn("-o", cmd)
        self.assertIn(str(Path("output.docx").absolute()), cmd)

        # The bundled reference doc must be passed so pandoc emits APA-styled
        # headings (Times New Roman, automatic color) instead of the template
        # defaults.
        ref_args = [c for c in cmd if c.startswith("--reference-doc=")]
        self.assertEqual(len(ref_args), 1, f"Expected --reference-doc in pandoc cmd: {cmd}")
        ref_path = ref_args[0].split("=", 1)[1]
        self.assertTrue(ref_path.endswith("pandoc_reference.docx"))
        self.assertTrue(Path(ref_path).is_file(), f"Reference doc missing: {ref_path}")


class TestPandocReferenceDoc(unittest.TestCase):
    """The bundled pandoc reference must carry APA-compliant heading styles.

    Pandoc copies the heading styles from the reference into every DOCX, so
    the template must ship Times New Roman, an 'automatic' (theme-resolving)
    color, and no theme-font/theme-color slots — otherwise Word would render
    the template defaults on headings.
    """

    _NS_MAIN = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def test_reference_doc_exists(self):
        self.assertTrue(REFERENCE_DOCX.is_file(), f"Missing {REFERENCE_DOCX}")

    def test_heading_styles_times_new_roman_automatic_no_theme(self):
        doc = Document(str(REFERENCE_DOCX))
        automatic_names = (
            "Heading 1",
            "Heading 2",
            "Heading 3",
            "Heading 4",
            "Heading 5",
        )
        for name in (*automatic_names, "TOC Heading"):
            style = doc.styles[name]
            self.assertEqual(
                style.font.name,
                "Times New Roman",
                f"{name} should use Times New Roman",
            )
            if name in automatic_names:
                color = style.font.color
                self.assertEqual(
                    color.type,
                    MSO_COLOR_TYPE.AUTO,
                    f"{name} color should be automatic, got {color.type!r}",
                )
                self.assertIsNone(color.rgb, f"{name} should not carry an explicit RGB color")
            r_fonts = style.element.find(f"{{{self._NS_MAIN}}}rPr/{{{self._NS_MAIN}}}rFonts")
            self.assertIsNotNone(r_fonts, f"{name} should have an rFonts element")
            for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
                self.assertNotIn(
                    f"{{{self._NS_MAIN}}}{attr}",
                    r_fonts.attrib,
                    f"{name} should have no {attr} in the reference doc",
                )

    @patch("subprocess.run")
    def test_run_failure_pandoc_error(self, mock_run):
        mock_result = MagicMock()
        mock_result.returncode = 1
        mock_result.stderr = "Error converting"
        mock_run.return_value = mock_result

        runner = PandocRunner()
        success = runner.run("# Title", "output.docx")

        self.assertFalse(success)

    @patch("subprocess.run")
    def test_run_pandoc_not_found(self, mock_run):
        mock_run.side_effect = FileNotFoundError("pandoc not found")

        runner = PandocRunner()
        success = runner.run("# Title", "output.docx")

        self.assertFalse(success)


if __name__ == "__main__":
    unittest.main()
