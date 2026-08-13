"""PyInstaller freeze entrypoint for the standalone xgen-seepage connector.

Kept as a tiny separate launcher script (rather than freezing the package's
`-m` module directly) because PyInstaller's Analysis needs a concrete .py
file as its entry point.
"""
from xgen_seepage.connector.app import main

if __name__ == "__main__":
    raise SystemExit(main())
