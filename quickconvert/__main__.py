"""Entry point used by the right-click menu: pythonw run.pyw <target> <file>..."""

import sys
from pathlib import Path

from .batch import Batch
from .common import show_error


def convert(source: Path, target: str) -> Path:
    # Imported here, not at the top: with a multi-select, Explorer starts one process per file and
    # all but one only queue their file, so they shouldn't pay for loading Pillow / pywin32.
    from . import images, markdown_pdf, office

    ext = source.suffix.lower()
    if target in ("pdf", "docx") and ext in markdown_pdf.SOURCE_EXTS:
        return markdown_pdf.convert(source, target)
    if target == "pdf" and ext in office.SOURCE_EXTS:
        return office.convert(source)
    if ext in images.SOURCE_EXTS and target in images.TARGETS:
        return images.convert(source, target)
    raise ValueError(f"Don't know how to convert {ext} to {target}.")


def run_without_window(batch: Batch, target: str) -> int:
    """Fallback when tkinter is missing: convert silently, pop up only on errors."""
    batch.wait_for_selection()
    failures = []
    while (source := batch.next_file()) is not None:
        try:
            convert(source, target)
        except Exception as exc:
            failures.append(f"{source.name}: {exc}")
    if failures:
        show_error("Could not convert:\n\n" + "\n".join(failures))
    return 1 if failures else 0


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        show_error("Usage: quickconvert <target> <file>...")
        return 2
    target = argv[0].lower()
    batch = Batch(target)
    if not batch.submit([Path(f) for f in argv[1:]]):
        return 0  # another QuickConvert process is already converting this selection
    try:
        from .window import ProgressWindow
    except ImportError:
        return run_without_window(batch, target)
    return ProgressWindow(batch, target, convert).run()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
