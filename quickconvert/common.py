import ctypes
from pathlib import Path


def output_path(source: Path, ext: str) -> Path:
    """Next to the source, same name, new extension; never overwrite ("name (1).ext")."""
    out = source.with_suffix(ext)
    n = 1
    while out.exists():
        out = source.with_name(f"{source.stem} ({n}){ext}")
        n += 1
    return out


def show_error(message: str) -> None:
    MB_ICONERROR = 0x10
    ctypes.windll.user32.MessageBoxW(None, message, "QuickConvert", MB_ICONERROR)
