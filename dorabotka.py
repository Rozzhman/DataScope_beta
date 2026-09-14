import os
import sys
from pathlib import Path


def _configure_tk_env() -> None:
    if not getattr(sys, "frozen", False):
        return
    base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    tcl_dir = base / "tcl" / "tcl8.6"
    tk_dir = base / "tcl" / "tk8.6"
    if tcl_dir.exists():
        os.environ.setdefault("TCL_LIBRARY", str(tcl_dir))
    if tk_dir.exists():
        os.environ.setdefault("TK_LIBRARY", str(tk_dir))


_configure_tk_env()

from legacy_api import *  # backward compatibility shim


if __name__ == "__main__":
    from absolute_v5.ui.main_window import App

    App().mainloop()
