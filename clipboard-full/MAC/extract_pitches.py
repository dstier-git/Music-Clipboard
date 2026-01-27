from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.extract_pitches import extract_pitches_from_mscx, main

__all__ = ["extract_pitches_from_mscx", "main"]

if __name__ == "__main__":
    main()
