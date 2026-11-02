"""Add (or with --uninstall, remove) the QuickConvert right-click menus for the current user.

Everything goes under HKEY_CURRENT_USER, so no admin rights are needed.
"""

import ctypes
import sys
import winreg
from pathlib import Path

from quickconvert import images, markdown_pdf, office

ROOT = r"Software\Classes\SystemFileAssociations"
KEY_PREFIX = "QuickConvert"
ICON = r"%SystemRoot%\System32\shell32.dll,-16769"
LABELS = {"pdf": "PDF", "docx": "Word (.docx)", "jpg": "JPG", "png": "PNG", "webp": "WebP"}
ALL_EXTS = markdown_pdf.SOURCE_EXTS + office.SOURCE_EXTS + images.SOURCE_EXTS


def targets_for(ext: str) -> list[str]:
    """Formats offered in the "Convert to" submenu for one file extension."""
    if ext in markdown_pdf.SOURCE_EXTS:
        return ["pdf", "docx"]
    if ext in office.SOURCE_EXTS:
        return ["pdf"]
    own = "jpg" if ext in images.JPEG_EXTS else ext.lstrip(".")
    return [t for t in ("jpg", "png", "webp", "pdf") if t != own]


def command(target: str) -> str:
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    launcher = Path(__file__).resolve().with_name("run.pyw")
    return f'"{pythonw}" "{launcher}" {target} "%1"'


def set_key(path: str, values: dict[str, str]) -> None:
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, path) as key:
        for name, value in values.items():
            kind = winreg.REG_EXPAND_SZ if "%" in value and name == "Icon" else winreg.REG_SZ
            winreg.SetValueEx(key, name or None, 0, kind, value)


def delete_tree(path: str) -> None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                delete_tree(f"{path}\\{child}")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
    except FileNotFoundError:
        pass


def add_menu(ext: str) -> None:
    menu = rf"{ROOT}\{ext}\shell\{KEY_PREFIX}"
    set_key(menu, {"MUIVerb": "Convert to", "Icon": ICON, "SubCommands": ""})
    for i, target in enumerate(targets_for(ext)):
        item = rf"{menu}\shell\{i}{target}"  # numeric prefix keeps menu order
        set_key(item, {"MUIVerb": LABELS[target]})
        set_key(rf"{item}\command", {"": command(target)})


def uninstall() -> None:
    for ext in ALL_EXTS:
        # .PDF / .Word are entries from older versions.
        for name in (KEY_PREFIX, f"{KEY_PREFIX}.PDF", f"{KEY_PREFIX}.Word"):
            delete_tree(rf"{ROOT}\{ext}\shell\{name}")


def install() -> None:
    uninstall()  # start clean so re-running picks up a moved folder or new Python
    for ext in ALL_EXTS:
        add_menu(ext)


def refresh_explorer() -> None:
    SHCNE_ASSOCCHANGED = 0x08000000
    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, 0, None, None)


if __name__ == "__main__":
    if "--uninstall" in sys.argv:
        uninstall()
        print("QuickConvert removed from the right-click menu.")
    else:
        install()
        print("QuickConvert installed. Right-click a file (Windows 11: 'Show more options').")
    refresh_explorer()
