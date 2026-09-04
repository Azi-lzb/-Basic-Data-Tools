"""Windows single-instance guard shared by every packaged edition."""

from __future__ import annotations

import atexit
import ctypes
import sys
from typing import Any


_MUTEX_NAME = "Local\\BaseDataAuditTool_SingleInstance_v1"
_ERROR_ALREADY_EXISTS = 183
_handle: Any | None = None


def acquire() -> bool:
    """Return ``True`` only for the first tool process in this user session."""
    global _handle
    if sys.platform != "win32":
        return True
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    handle = kernel32.CreateMutexW(None, False, _MUTEX_NAME)
    if not handle:
        # Failing open is safer than blocking an audit merely because a locked
        # down endpoint prevents named mutex creation.
        return True
    if kernel32.GetLastError() == _ERROR_ALREADY_EXISTS:
        kernel32.CloseHandle(ctypes.c_void_p(handle))
        return False
    _handle = handle
    atexit.register(_release)
    return True


def show_already_running_message() -> None:
    """Show a short desktop message without depending on tkinter or the bridge."""
    if sys.platform == "win32":
        ctypes.windll.user32.MessageBoxW(
            None,
            "基础数据审核工具已在运行。请切换到已打开的窗口，完成后再启动新的任务。",
            "基础数据审核工具",
            0x30,
        )


def _release() -> None:
    global _handle
    if _handle and sys.platform == "win32":
        ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(_handle))
    _handle = None
