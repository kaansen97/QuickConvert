import html
import tempfile
from pathlib import Path

import markdown2

from .common import output_path
from .office import export_docx, export_pdf

SOURCE_EXTS = [".md", ".markdown"]

EXTRAS = ["tables", "fenced-code-blocks", "header-ids", "strike", "task_list", "cuddled-lists"]

# Word's HTML import understands a subset of CSS; keep it simple.
CSS = """
body { font-family: "Segoe UI", Arial, sans-serif; font-size: 10.5pt; line-height: 1.4;
       color: #1f2328; }
h1 { font-size: 22pt; border-bottom: 1px solid #d0d7de; }
h2 { font-size: 16pt; border-bottom: 1px solid #d0d7de; margin-top: 18pt; }
h3 { font-size: 13pt; margin-top: 14pt; }
a { color: #0969da; text-decoration: none; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #d0d7de; padding: 4px 8px; text-align: left; vertical-align: top; }
th { background: #f6f8fa; }
code { font-family: Consolas, monospace; font-size: 9.5pt; background: #f2f4f6; }
pre { background: #f6f8fa; padding: 8px; font-family: Consolas, monospace; font-size: 9.5pt; }
blockquote { margin-left: 0; padding-left: 10pt; color: #59636e; border-left: 3pt solid #d0d7de; }
"""


def to_html(source: Path) -> str:
    body = markdown2.markdown(source.read_text(encoding="utf-8-sig"), extras=EXTRAS)
    # <base> makes relative links such as images/foo.svg resolve against the .md's folder.
    base = source.parent.resolve().as_uri() + "/"
    return (
        "<!doctype html><html><head><meta charset='utf-8'>"
        f"<base href='{html.escape(base)}'><title>{html.escape(source.stem)}</title>"
        f"<style>{CSS}</style></head><body>{body}</body></html>"
    )


def convert(source: Path, target: str = "pdf") -> Path:
    """Markdown -> HTML -> PDF or .docx via a hidden Word (Edge headless is blocked by policy here)."""
    out = output_path(source, ".docx" if target == "docx" else ".pdf")
    with tempfile.TemporaryDirectory(prefix="quickconvert-") as tmp:
        page = Path(tmp) / "page.html"
        page.write_text(to_html(source), encoding="utf-8")
        if target == "docx":
            export_docx(page, out)
        else:
            export_pdf(page, out, fit_images=True)
    return out
