"""Entry point used by the right-click menu: pythonw -m quickconvert <target> <file>..."""

import sys
import traceback
from pathlib import Path

from . import images, markdown_pdf, office
from .common import show_error


def convert(source: Path, target: str) -> Path:
    ext = source.suffix.lower()
    if target in ("pdf", "docx") and ext in markdown_pdf.SOURCE_EXTS:
        return markdown_pdf.convert(source, target)
    if target == "pdf" and ext in office.SOURCE_EXTS:
        return office.convert(source)
    if ext in images.SOURCE_EXTS and target in images.TARGETS:
        return images.convert(source, target)
    raise ValueError(f"Don't know how to convert {ext} to {target}.")


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        show_error("Usage: quickconvert <target> <file>...")
        return 2
    target, files = argv[0].lower(), argv[1:]
    failures = []
    for name in files:
        try:
            convert(Path(name), target)
        except Exception as exc:
            traceback.print_exc()
            failures.append(f"{Path(name).name}: {exc}")
    if failures:
        show_error("Could not convert:\n\n" + "\n".join(failures))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
