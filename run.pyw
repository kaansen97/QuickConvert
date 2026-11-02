# Launcher the right-click menu calls with pythonw (no console window).
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from quickconvert.__main__ import main  # noqa: E402

sys.exit(main(sys.argv[1:]))
