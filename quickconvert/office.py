"""Conversions done by a hidden Microsoft Office (Word, Excel, PowerPoint) over COM."""

from contextlib import contextmanager
from pathlib import Path

import pythoncom
import win32com.client

from .common import output_path

WORD_EXTS = [".docx", ".doc", ".docm", ".rtf", ".odt"]
EXCEL_EXTS = [".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv"]
POWERPOINT_EXTS = [".pptx", ".ppt", ".pptm", ".odp"]
SVG_EXTS = [".svg"]
SOURCE_EXTS = WORD_EXTS + EXCEL_EXTS + POWERPOINT_EXTS + SVG_EXTS

WD_FORMAT_PDF = 17
WD_FORMAT_DOCX = 16
XL_TYPE_PDF = 0
PP_SAVE_AS_PDF = 32


@contextmanager
def office_app(prog_id: str):
    pythoncom.CoInitialize()
    # DispatchEx starts a private, hidden instance so files you already have open are untouched.
    app = win32com.client.DispatchEx(prog_id)
    try:
        yield app
    finally:
        # PowerPoint is single-instance: only quit if we are not sharing it with your own work.
        if prog_id != "PowerPoint.Application" or app.Presentations.Count == 0:
            app.Quit()
        pythoncom.CoUninitialize()


@contextmanager
def word_document(path: Path):
    """Open any file Word understands (doc, docx, odt, rtf, html) read-only."""
    with office_app("Word.Application") as word:
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            str(path), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
        )
        try:
            yield doc
        finally:
            doc.Close(SaveChanges=False)


def export_pdf(document: Path, out: Path, fit_images: bool = False) -> None:
    with word_document(document) as doc:
        if fit_images:
            _shrink_wide_images(doc)
        doc.ExportAsFixedFormat(str(out), WD_FORMAT_PDF)


def export_docx(document: Path, out: Path) -> None:
    """Save an HTML page as a self-contained .docx."""
    with word_document(document) as doc:
        _shrink_wide_images(doc)
        _embed_linked_images(doc)
        doc.SaveAs2(str(out), FileFormat=WD_FORMAT_DOCX)


def _shrink_wide_images(doc) -> None:
    setup = doc.PageSetup
    usable_w = setup.PageWidth - setup.LeftMargin - setup.RightMargin
    usable_h = setup.PageHeight - setup.TopMargin - setup.BottomMargin
    for shape in doc.InlineShapes:
        ratio = min(usable_w / shape.Width, usable_h / shape.Height, 1)
        if ratio < 1:
            shape.Width, shape.Height = shape.Width * ratio, shape.Height * ratio


def _embed_linked_images(doc) -> None:
    """HTML images arrive as links to files on disk; store them inside the .docx instead."""
    for shape in doc.InlineShapes:
        link = shape.LinkFormat
        if link is not None:
            link.SavePictureWithDocument = True
            link.BreakLink()


def _svg_to_pdf(source: Path, out: Path) -> None:
    with office_app("Word.Application") as word:
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Add()
        try:
            doc.InlineShapes.AddPicture(str(source), LinkToFile=False, SaveWithDocument=True)
            _shrink_wide_images(doc)
            doc.ExportAsFixedFormat(str(out), WD_FORMAT_PDF)
        finally:
            doc.Close(SaveChanges=False)


def _excel_to_pdf(source: Path, out: Path) -> None:
    with office_app("Excel.Application") as excel:
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        try:
            book.ExportAsFixedFormat(XL_TYPE_PDF, str(out))
        finally:
            book.Close(SaveChanges=False)


def _powerpoint_to_pdf(source: Path, out: Path) -> None:
    with office_app("PowerPoint.Application") as powerpoint:
        deck = powerpoint.Presentations.Open(
            str(source), ReadOnly=True, Untitled=False, WithWindow=False
        )
        try:
            deck.SaveAs(str(out), PP_SAVE_AS_PDF)
        finally:
            deck.Close()


def convert(source: Path) -> Path:
    out = output_path(source, ".pdf")
    ext = source.suffix.lower()
    if ext in EXCEL_EXTS:
        _excel_to_pdf(source, out)
    elif ext in POWERPOINT_EXTS:
        _powerpoint_to_pdf(source, out)
    elif ext in SVG_EXTS:
        _svg_to_pdf(source, out)
    else:
        export_pdf(source, out)
    return out
