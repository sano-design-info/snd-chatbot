import base64
from pathlib import Path

# dirpath
ROOTDIR = Path(__file__).parents[1]
EXPORTDIR_PATH = ROOTDIR / "exportdir"
EXPORTDIR_PATH.mkdir(parents=True, exist_ok=True)


def decode_base64url(s) -> bytes:
    return base64.urlsafe_b64decode(s) + b"=" * (4 - (len(s) % 4))
