"""Write the public GitHub repository into the build resource, not user config."""

from __future__ import annotations

import json
import sys
from pathlib import Path


if len(sys.argv) != 2 or sys.argv[1].count("/") != 1:
    raise SystemExit("Usage: set_release_repository.py owner/repository")

Path("assets/update_source.json").write_text(
    json.dumps({"repository": sys.argv[1]}, ensure_ascii=False, indent=2) + "\n",
    encoding="utf-8",
)
