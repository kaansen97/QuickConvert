# QuickConvert

Right-click file conversion for Windows, in the spirit of [FileConverter](https://github.com/Tichau/FileConverter),
but **no admin rights needed**: everything installs for the current user only.

Every supported file gets one **Convert to ▸** entry; pick the format from the submenu.

| Right-click a…                                  | Convert to ▸                          | Result                       |
|-------------------------------------------------|---------------------------------------|------------------------------|
| `.md` / `.markdown`                             | **PDF / Word (.docx)**                | Styled document, tables and images (incl. SVG) included; images are embedded in the .docx |
| `.docx` `.doc` `.docm` `.rtf` `.odt`            | **PDF**                               | Exported by a hidden Word, Word never opens |
| `.xlsx` `.xls` `.xlsm` `.xlsb` `.ods` `.csv`    | **PDF**                               | Exported by a hidden Excel, using each sheet's print settings |
| `.pptx` `.ppt` `.pptm` `.odp`                   | **PDF**                               | Exported by a hidden PowerPoint, one slide per page |
| `.svg`                                          | **PDF**                               | Drawn by Word, scaled to fit the page |
| `.jpg` `.jpeg` `.jfif` `.png` `.webp` `.bmp` `.tif(f)` `.ico` `.dds` `.tga` `.psd` | **JPG / PNG / WebP / PDF** | Same image, new format |

The new file is saved next to the original with the same name. Nothing is overwritten:
if `Report.pdf` exists you get `Report (1).pdf`. Select several files to convert them all.

> Windows 11: the entries are under **Show more options** (or Shift + right-click).

## Requirements

- Windows 10/11
- Python 3.10+ ([python.org](https://www.python.org/downloads/) — the per-user installer needs no admin)
- Microsoft Office: Word (Word, Markdown, SVG), Excel and PowerPoint for their own files

## Install

1. Put this folder somewhere permanent (the menu points at it).
2. Double-click **`install.bat`**. It installs the Python packages for your user and adds the menu entries.

Moved the folder or changed Python? Run `install.bat` again.

## Uninstall

Double-click **`uninstall.bat`**.

## How it works

- `install.py` writes menu entries under `HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\<ext>\shell`.
  Each entry runs `pythonw.exe run.pyw <target> "<file>"` (no console window).
- **Markdown → PDF:** `markdown2` turns the file into HTML, then a hidden Word instance opens the
  HTML, shrinks over-wide images to the page and exports a PDF. (Edge headless printing was the first
  choice but is blocked by corporate Edge policies.)
- **Office → PDF:** a private hidden Word / Excel / PowerPoint (`DispatchEx`) opens the file read-only and
  exports it. Files you already have open are not touched; PowerPoint is never closed if you are using it.
- **Images:** Pillow. EXIF rotation is applied; transparency becomes white for JPG and PDF.
  For `.psd` the flattened composite image is used.
- Not supported (on purpose, to avoid extra dependencies): HEIC, camera RAW, EXR, audio/video.
- Errors show a small popup with the file name and the reason.

## Development

```bash
py -m unittest discover tests -v
```

## License

MIT
