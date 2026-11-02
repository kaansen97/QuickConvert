from pathlib import Path

from PIL import Image, ImageOps

from .common import output_path, progress

# target -> (extension, Pillow format)
TARGETS = {
    "jpg": (".jpg", "JPEG"),
    "png": (".png", "PNG"),
    "webp": (".webp", "WEBP"),
    "pdf": (".pdf", "PDF"),
}

JPEG_EXTS = [".jpg", ".jpeg", ".jfif"]
SOURCE_EXTS = JPEG_EXTS + [".png", ".webp", ".bmp", ".tif", ".tiff", ".ico", ".dds", ".tga", ".psd"]


def _flatten(img: Image.Image) -> Image.Image:
    """Drop transparency onto a white background (JPG and PDF have no alpha)."""
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        background = Image.new("RGB", img.size, "white")
        background.paste(img, mask=img.getchannel("A"))
        return background
    return img.convert("RGB")


def convert(source: Path, target: str) -> Path:
    ext, fmt = TARGETS[target]
    out = output_path(source, ext)
    progress("Reading image…", 0.2)
    with Image.open(source) as img:
        img = ImageOps.exif_transpose(img)
        if fmt in ("JPEG", "PDF"):
            img = _flatten(img)
        options = {"quality": 92} if fmt in ("JPEG", "WEBP") else {}
        progress(f"Saving {target.upper()}…", 0.6)
        img.save(out, fmt, **options)
    return out
