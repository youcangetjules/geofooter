"""Attach to the Outlook the user has open. Never start one.

``Dispatch("Outlook.Application")`` on a closed Outlook starts a headless
``OUTLOOK.EXE -Embedding`` that exits when COM is released. Add-ins and VBA
startup then run with no window, and the add-in DLL stays locked (BUG-040).
Every Python caller in the suite goes through ``attach_outlook`` instead.
"""

from __future__ import annotations

import ctypes
import logging
from typing import Any

logger = logging.getLogger(__name__)

# Explorer and inspector windows both use this class. A headless Outlook owns none.
_OUTLOOK_WINDOW_CLASS = "rctrl_renwnd32"


def outlook_is_running() -> bool:
    """True when Outlook has a window (shown, minimised or hidden to the tray).

    A window lookup, not a process list: it costs microseconds, so it is safe on
    a UI thread, and it is False for a headless COM-only Outlook.
    """
    try:
        find = ctypes.windll.user32.FindWindowW
        find.restype = ctypes.c_void_p
        return bool(find(_OUTLOOK_WINDOW_CLASS, None))
    except Exception as exc:
        logger.debug("FindWindowW failed: %s", exc)
        return False


def attach_outlook() -> Any:
    """Return the running Outlook.Application, or raise RuntimeError if it is closed."""
    import win32com.client  # type: ignore

    try:
        return win32com.client.GetObject(Class="Outlook.Application")
    except Exception:
        # GetObject can miss a running Outlook; Dispatch then attaches to it.
        if not outlook_is_running():
            raise RuntimeError("Outlook is not open") from None
        return win32com.client.Dispatch("Outlook.Application")
