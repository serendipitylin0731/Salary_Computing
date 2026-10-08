from __future__ import annotations

import ctypes
import sys


# Windows 10 1703+：窗口移动到不同缩放比例的显示器时也能保持清晰。
_DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = ctypes.c_void_p(-4)


def enable_high_dpi_awareness() -> bool:
    """在创建任何 Tk 窗口前启用 Windows 高 DPI 感知。"""
    if sys.platform != "win32":
        return False

    try:
        user32 = ctypes.windll.user32
        set_process_context = getattr(user32, "SetProcessDpiAwarenessContext", None)
        if set_process_context and set_process_context(
            _DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
        ):
            return True

        # 若宿主进程已经设置过 DPI 模式，进程级调用会被拒绝；此时为
        # 创建 Tk 和系统文件对话框的当前线程设置 Per-Monitor V2。
        set_thread_context = getattr(user32, "SetThreadDpiAwarenessContext", None)
        if set_thread_context and set_thread_context(
            _DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
        ):
            return True
    except (AttributeError, OSError):
        pass

    try:
        # Windows 8.1+ 的回退方案：PROCESS_PER_MONITOR_DPI_AWARE。
        if ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0:
            return True
    except (AttributeError, OSError):
        pass

    try:
        # Windows Vista/7 的回退方案：系统 DPI 感知。
        return bool(ctypes.windll.user32.SetProcessDPIAware())
    except (AttributeError, OSError):
        return False
