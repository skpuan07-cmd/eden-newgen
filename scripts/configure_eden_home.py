#!/usr/bin/env python3
"""Configure the local web2py welcome page to open Eden."""

import sys
from pathlib import Path


ORIGINAL_INDEX = '''def index():
    response.flash = T("Hello World")
    return dict(message=T('Welcome to web2py!'))
'''
EDEN_INDEX = '''def index():
    redirect(URL('eden', 'default', 'index'))
'''


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit("Usage: configure_eden_home.py WEB2PY_DIR")

    controller = (
        Path(sys.argv[1])
        / "applications"
        / "welcome"
        / "controllers"
        / "default.py"
    )
    source = controller.read_text(encoding="utf-8")

    if EDEN_INDEX in source:
        return
    if ORIGINAL_INDEX not in source:
        raise SystemExit(
            f"Cannot configure Eden home: expected welcome index not found in {controller}"
        )

    controller.write_text(
        source.replace(ORIGINAL_INDEX, EDEN_INDEX, 1),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
