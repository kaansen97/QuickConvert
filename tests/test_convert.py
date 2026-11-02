"""Converts real sample files. Run with: py -m unittest discover tests"""

import shutil
import tempfile
import unittest
from pathlib import Path

from PIL import Image

from quickconvert.__main__ import convert
from quickconvert.common import output_path

SVG = """<svg xmlns="http://www.w3.org/2000/svg" width="200" height="80">
<rect width="200" height="80" fill="#4a90d9"/><text x="20" y="48" fill="white">diagram</text></svg>"""

MARKDOWN = """# Sample manual

Some **bold** text and a table:

| Step | What happens |
|------|--------------|
| 1    | Draft        |
| 2    | Publish      |

---

![diagram](images/diagram.svg)
"""


def word_available() -> bool:
    try:
        import win32com.client

        win32com.client.DispatchEx("Word.Application").Quit()
        return True
    except Exception:
        return False


class ConvertTests(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def assert_pdf(self, path: Path):
        self.assertTrue(path.exists(), path)
        self.assertEqual(path.read_bytes()[:5], b"%PDF-")

    def test_output_path_never_overwrites(self):
        src = self.dir / "a.md"
        src.touch()
        (self.dir / "a.pdf").touch()
        self.assertEqual(output_path(src, ".pdf").name, "a (1).pdf")

    def test_transparent_png_to_jpg_webp_pdf(self):
        src = self.dir / "logo.png"
        Image.new("RGBA", (40, 30), (255, 0, 0, 0)).save(src)
        jpg = convert(src, "jpg")
        with Image.open(jpg) as img:
            self.assertEqual((img.format, img.size), ("JPEG", (40, 30)))
            self.assertEqual(img.getpixel((5, 5)), (255, 255, 255))  # alpha flattened to white
        with Image.open(convert(src, "webp")) as img:
            self.assertEqual(img.format, "WEBP")
        self.assert_pdf(convert(src, "pdf"))

    def test_jpeg_to_png(self):
        src = self.dir / "photo.jpeg"
        Image.new("RGB", (10, 10), "blue").save(src)
        with Image.open(convert(src, "png")) as img:
            self.assertEqual(img.format, "PNG")

    def test_ico_and_tga_to_png(self):
        for name in ("icon.ico", "texture.tga"):
            src = self.dir / name
            Image.new("RGBA", (32, 32), (0, 200, 0, 255)).save(src)
            with Image.open(convert(src, "png")) as img:
                self.assertEqual(img.format, "PNG")

    @unittest.skipUnless(word_available(), "Word not installed")
    def test_svg_to_pdf(self):
        src = self.dir / "diagram.svg"
        src.write_text(SVG)
        self.assert_pdf(convert(src, "pdf"))

    def test_excel_to_pdf(self):
        try:
            import win32com.client

            win32com.client.DispatchEx("Excel.Application").Quit()
        except Exception:
            self.skipTest("Excel not installed")
        import openpyxl

        src = self.dir / "Budget.xlsx"
        book = openpyxl.Workbook()
        book.active.append(["Item", "Cost"])
        book.active.append(["Laptop", 1200])
        book.save(src)
        self.assert_pdf(convert(src, "pdf"))

    def test_powerpoint_to_pdf(self):
        try:
            import pptx
            import win32com.client

            app = win32com.client.DispatchEx("PowerPoint.Application")
            if app.Presentations.Count == 0:  # don't close a PowerPoint you are using
                app.Quit()
        except Exception:
            self.skipTest("PowerPoint or python-pptx not installed")
        src = self.dir / "Deck.pptx"
        deck = pptx.Presentation()
        deck.slides.add_slide(deck.slide_layouts[0]).shapes.title.text = "Hello"
        deck.save(src)
        self.assert_pdf(convert(src, "pdf"))

    def test_unknown_conversion_raises(self):
        src = self.dir / "notes.txt"
        src.write_text("hi")
        with self.assertRaises(ValueError):
            convert(src, "pdf")

    @unittest.skipUnless(word_available(), "Word not installed")
    def test_markdown_to_pdf_with_image(self):
        (self.dir / "images").mkdir()
        (self.dir / "images" / "diagram.svg").write_text(SVG)
        src = self.dir / "Sample Manual.md"
        src.write_text(MARKDOWN, encoding="utf-8")
        out = convert(src, "pdf")
        self.assert_pdf(out)
        self.assertEqual(out.name, "Sample Manual.pdf")

    @unittest.skipUnless(word_available(), "Word not installed")
    def test_markdown_emoji_keep_colour(self):
        try:
            import pymupdf
        except ImportError:
            self.skipTest("pymupdf not installed")
        src = self.dir / "Colours.md"
        src.write_text("| Colour | Name |\n|---|---|\n| \U0001F7E8 Yellow | Pending |\n\n"
                       "> ✅ **Good to know**\n", encoding="utf-8")
        pix = pymupdf.open(convert(src, "pdf"))[0].get_pixmap(dpi=40)
        colourful = sum(
            1 for y in range(pix.height) for x in range(pix.width)
            if max(pix.pixel(x, y)[:3]) - min(pix.pixel(x, y)[:3]) > 60
        )
        self.assertGreater(colourful, 20)  # Word draws emoji black-and-white when text has a colour

    @unittest.skipUnless(word_available(), "Word not installed")
    def test_markdown_to_docx_embeds_image(self):
        import docx

        (self.dir / "images").mkdir()
        (self.dir / "images" / "diagram.svg").write_text(SVG)
        src = self.dir / "Sample Manual.md"
        src.write_text(MARKDOWN, encoding="utf-8")
        out = convert(src, "docx")
        self.assertEqual(out.name, "Sample Manual.docx")
        shutil.rmtree(self.dir / "images")  # image must live inside the .docx, not on disk
        document = docx.Document(out)
        self.assertEqual(len(document.tables), 1)
        self.assertTrue(document.inline_shapes)
        self.assertTrue(any(r.reltype.endswith("/image") and not r.is_external
                            for r in document.part.rels.values()))

    @unittest.skipUnless(word_available(), "Word not installed")
    def test_word_to_pdf(self):
        import docx

        src = self.dir / "Report.docx"
        document = docx.Document()
        document.add_heading("Report", 0)
        document.add_paragraph("Hello from QuickConvert.")
        document.save(src)
        self.assert_pdf(convert(src, "pdf"))


if __name__ == "__main__":
    unittest.main()
