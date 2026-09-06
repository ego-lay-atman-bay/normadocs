"""Regenerate the bundled pandoc reference docx with APA-styled defaults.

Pandoc copies the styles from its reference document into every DOCX it
writes (``--reference-doc``). Pandoc's built-in template ships Heading 1-5
with an accent theme color and theme-font slots (``w:asciiTheme`` etc.),
which makes the formatters' post-processing fight the SwP defaults for
font/color. Baking the APA defaults into the reference means headings are
Times New Roman, black from the start.

Run from the repo root:

    python scripts/generate_pandoc_reference_docx.py

Requires:
    - ``pandoc`` on PATH
    - python-docx and the normadocs package importable (used for the same
      APA style configuration the formatter applies at runtime, so the
      template and the formatting pipeline cannot drift apart).
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from normadocs.formatters.apa.apa_styles import APAStylesHandler

RESOURCE = SRC / "normadocs" / "resources" / "pandoc_reference.docx"
NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DEFAULT_BODY_FONT = "Times New Roman"
_THEME_FONT_ATTRS = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme")
_THEME_COLOR_ATTRS = ("themeColor", "themeShade", "themeTint")


def load_default_reference() -> bytes:
    """Return pandoc's built-in reference.docx bytes."""
    if shutil.which("pandoc") is None:
        sys.exit("pandoc is required to regenerate the reference docx")
    proc = subprocess.run(
        ["pandoc", "--print-default-data-file", "reference.docx"],
        capture_output=True,
        check=False,
    )
    if proc.returncode != 0:
        sys.exit(proc.stderr.decode() or "pandoc failed to print the default reference.docx")
    if not proc.stdout:
        sys.exit("pandoc returned an empty default reference.docx")
    return proc.stdout


def harden_styles(doc: Document) -> None:
    """Force Times New Roman + black on every heading and strip theme fonts.

    ``create_styles`` already handles Heading 1-5, Normal, Body Text, and
    Compact. This pass covers the remaining styles in pandoc's template
    (Heading 6-9, TOC Heading) and removes any leftover theme-font or
    theme-color attributes that Word would otherwise resolve first.
    """
    for style_el in doc.styles.element.iter(f"{{{NS}}}style"):
        name_el = style_el.find(f"{{{NS}}}name")
        name = name_el.get(f"{{{NS}}}val") if name_el is not None else ""
        r_pr = style_el.find(f"{{{NS}}}rPr")
        if r_pr is None:
            continue
        r_fonts = r_pr.find(f"{{{NS}}}rFonts")
        if r_fonts is not None:
            for attr in _THEME_FONT_ATTRS:
                r_fonts.attrib.pop(f"{{{NS}}}{attr}", None)
        if not (name.lower().startswith("heading") or name.lower() == "toc heading"):
            continue
        if r_fonts is not None:
            r_fonts.set(f"{{{NS}}}ascii", DEFAULT_BODY_FONT)
            r_fonts.set(f"{{{NS}}}hAnsi", DEFAULT_BODY_FONT)
        color = r_pr.find(f"{{{NS}}}color")
        if color is None:
            color = etree.SubElement(r_pr, f"{{{NS}}}color")
        color.set(f"{{{NS}}}val", "000000")
        for attr in _THEME_COLOR_ATTRS:
            color.attrib.pop(f"{{{NS}}}{attr}", None)


def build() -> None:
    """Generate the resource docx from pandoc's default with APA defaults."""
    default_reference = load_default_reference()

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(default_reference)
        tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    APAStylesHandler(doc).create_styles()
    harden_styles(doc)

    RESOURCE.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(RESOURCE))
    print(f"Wrote {RESOURCE}")


def summarize() -> None:
    """Print the resulting heading font/color for a quick sanity check."""
    doc = Document(str(RESOURCE))
    for i in range(1, 6):
        style = doc.styles[f"Heading {i}"]
        print(
            f"Heading {i}: font={style.font.name!r}, "
            f"color={style.font.color.rgb if style.font.color and style.font.color.type is not None else None}, "
            f"theme={style.font.color.theme_color if style.font.color else None}"
        )


if __name__ == "__main__":
    build()
    summarize()