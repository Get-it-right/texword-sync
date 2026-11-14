"""Focused end-to-end checks for generated Word programs and template styles."""
from __future__ import annotations

import os
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import texword as tw
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt
from lxml import etree as E


PYTHON = Path(sys.executable)
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"


def _png_fixture() -> bytes:
    def chunk(name: bytes, payload: bytes) -> bytes:
        body = name + payload
        return struct.pack(">I", len(payload)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 6, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(b"\x00\x7f\x23\x9a\xff"))
        + chunk(b"IEND", b"")
    )


def _make_template(path: Path, *, body_style: tuple[str, int, float] | None = None) -> str | None:
    doc = Document()
    doc.add_paragraph("COVER UNCHANGED")
    doc.add_heading("Body starts", level=1)
    sample = doc.add_paragraph("Body sample")
    style_id = None
    if body_style is not None:
        font_name, font_size, indent_cm = body_style
        sample_style = doc.styles.add_style("Body sample", WD_STYLE_TYPE.PARAGRAPH)
        sample_style.base_style = doc.styles["Normal"]
        sample_style.font.name = font_name
        sample_style.font.size = Pt(font_size)
        sample_style.paragraph_format.first_line_indent = Cm(indent_cm)
        sample.style = sample_style
        style_id = sample_style.style_id
    doc.add_paragraph("TAIL UNCHANGED")
    doc.save(path)
    return style_id


def _adapt(template: Path, adapter: Path) -> dict:
    return tw.adapt(
        template,
        adapter,
        body_start="Body starts",
        body_end="TAIL UNCHANGED",
        samples={"body": 2},
    )


def _style_signature(path: Path, style_id: str) -> tuple[str | None, str | None, str | None]:
    with zipfile.ZipFile(path) as archive:
        styles = E.fromstring(archive.read("word/styles.xml"))
    style = next(s for s in styles.findall("w:style", NS) if s.get(W + "styleId") == style_id)
    fonts = style.find("w:rPr/w:rFonts", NS)
    size = style.find("w:rPr/w:sz", NS)
    indent = style.find("w:pPr/w:ind", NS)
    return (
        (fonts.get(W + "ascii") or fonts.get(W + "hAnsi")) if fonts is not None else None,
        size.get(W + "val") if size is not None else None,
        indent.get(W + "firstLine") if indent is not None else None,
    )


class GeneratedCodeExtraTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_side_by_side_images_share_one_caption_and_roundtrip(self) -> None:
        template = self.root / "template.docx"
        _make_template(template)
        adapter = self.root / "adapter"
        _adapt(template, adapter)
        image_bytes = _png_fixture()
        (self.root / "left.png").write_bytes(image_bytes)
        (self.root / "right.png").write_bytes(image_bytes)
        source = self.root / "pair.tex"
        source.write_text(r"""\documentclass{article}
\usepackage{graphicx}
\begin{document}
\section{Paired image}
See Figure~\ref{fig:pair}.
\begin{figure}[htbp]\centering
\includegraphics[width=3cm]{left.png} \includegraphics[width=3cm]{right.png}
\caption{Shared panel caption}\label{fig:pair}
\end{figure}
\end{document}
""", encoding="utf8")
        code = self.root / "build_pair.py"
        tw.compile_code(adapter, source, code, roundtrip=True)
        output = self.root / "pair.docx"
        subprocess.run([str(PYTHON), str(code), str(output)], check=True,
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        doc = Document(output)
        root = doc.element
        picture_paragraphs = root.xpath('.//w:p[w:r/w:drawing]')
        self.assertEqual(len(picture_paragraphs), 1)
        self.assertEqual(len(picture_paragraphs[0].xpath('.//w:drawing')), 2)
        seq = [n.text for n in root.xpath('.//w:instrText') if 'SEQ Figure' in (n.text or '')]
        self.assertEqual(len(seq), 1)
        self.assertEqual(len(root.xpath('.//w:bookmarkStart[@w:name="'+tw.bookmark_name('fig:pair')+'"]')), 1)
        caption = next(p for p in root.xpath('.//w:p') if 'Shared panel caption' in ''.join(p.xpath('.//w:t/text()')))
        for node in caption.xpath('.//w:t'):
            if node.text and 'Shared panel caption' in node.text:
                node.text = node.text.replace('Shared panel caption', 'Updated shared caption')
        doc.save(output)
        from roundtrip import sync_word
        sync_word(output, source, apply=True)
        updated = source.read_text(encoding='utf8')
        self.assertEqual(updated.count('\\includegraphics'), 2)
        self.assertEqual(updated.count('\\caption{'), 1)
        self.assertIn('Updated shared caption', updated)
        self.assertIn('\\label{fig:pair}', updated)

    def test_embedded_image_bytes_survive_and_program_runs_without_pandoc(self) -> None:
        template = self.root / "template.docx"
        _make_template(template)
        adapter = self.root / "adapter"
        _adapt(template, adapter)

        source = self.root / "image.tex"
        image = self.root / "pixel.png"
        image_bytes = _png_fixture()
        image.write_bytes(image_bytes)
        source.write_text(
            r"""\documentclass{article}
\usepackage{graphicx}
\begin{document}
\section{Image generation}
\begin{figure}[htbp]\centering
\includegraphics[width=2cm]{pixel.png}
\caption{Pixel fixture}\label{fig:pixel}
\end{figure}
\end{document}
""",
            encoding="utf8",
        )
        code = self.root / "build_image.py"
        tw.compile_code(adapter, source, code)

        asset_files = list((self.root / "build_image_assets").glob("image_*"))
        self.assertEqual(len(asset_files), 1)
        self.assertEqual(asset_files[0].read_bytes(), image_bytes)

        output = self.root / "image-output.docx"
        hidden_source = source.with_suffix(".hidden")
        hidden_image = image.with_suffix(".hidden")
        source.rename(hidden_source)
        image.rename(hidden_image)
        try:
            env = dict(os.environ, PANDOC="INTENTIONALLY_UNAVAILABLE")
            subprocess.run(
                [str(PYTHON), str(code), str(output)],
                check=True,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
        finally:
            hidden_source.rename(source)
            hidden_image.rename(image)

        report = tw.load(output.with_suffix(".validation.json"))
        self.assertFalse(report["generation_requires_pandoc"])
        with zipfile.ZipFile(output) as archive:
            media = [archive.read(name) for name in archive.namelist() if name.startswith("word/media/")]
        self.assertIn(image_bytes, media)

    def test_same_latex_uses_each_templates_distinct_native_body_style(self) -> None:
        source = self.root / "shared.tex"
        source.write_text(
            r"""\documentclass{article}
\begin{document}
\section{Shared heading}
This is an opening paragraph.

This is the same body paragraph.
\end{document}
""",
            encoding="utf8",
        )
        specs = [("Arial", 11, 0.5), ("Courier New", 15, 1.25)]
        outputs: list[tuple[Path, tuple[str | None, str | None, str | None]]] = []

        for index, spec in enumerate(specs):
            template = self.root / f"template-{index}.docx"
            source_style_id = _make_template(template, body_style=spec)
            adapter = self.root / f"adapter-{index}"
            _adapt(template, adapter)
            code = self.root / f"build-{index}.py"
            tw.compile_code(adapter, source, code)
            output = self.root / f"output-{index}.docx"
            env = dict(os.environ, PANDOC="INTENTIONALLY_UNAVAILABLE")
            subprocess.run(
                [str(PYTHON), str(code), str(output)],
                check=True,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )

            paragraph = next(p for p in Document(output).paragraphs if p.text == "This is the same body paragraph.")
            self.assertEqual(paragraph.style.style_id, "BodyText")
            outputs.append((output, _style_signature(output, "BodyText")))
            self.assertEqual(_style_signature(template, source_style_id), outputs[-1][1])

        self.assertEqual([sig[0] for _, sig in outputs], ["Arial", "Courier New"])
        self.assertEqual([sig[1] for _, sig in outputs], ["22", "30"])
        self.assertNotEqual(outputs[0][1], outputs[1][1])


if __name__ == "__main__":
    unittest.main(verbosity=2)
