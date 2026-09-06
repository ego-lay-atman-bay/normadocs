"""
Tests for Pandoc Client.
"""

import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from normadocs.pandoc_client import PandocRunner


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
        # headings (black, Times New Roman) instead of the template defaults.
        ref_args = [c for c in cmd if c.startswith("--reference-doc=")]
        self.assertEqual(len(ref_args), 1, f"Expected --reference-doc in pandoc cmd: {cmd}")
        ref_path = ref_args[0].split("=", 1)[1]
        self.assertTrue(ref_path.endswith("pandoc_reference.docx"))
        self.assertTrue(Path(ref_path).is_file(), f"Reference doc missing: {ref_path}")

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
