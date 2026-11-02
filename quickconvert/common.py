import ctypes
from pathlib import Path

_progress_hook = None


def set_progress_hook(hook) -> None:
    """hook(step: str, fraction: float) is called by converters; None turns reporting off."""
    global _progress_hook
    _progress_hook = hook


def progress(step: str, fraction: float) -> None:
    if _progress_hook:
        _progress_hook(step, fraction)


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
