from __future__ import annotations

import sys
from pathlib import Path

from src.windows import enable_high_dpi_awareness


def application_root() -> Path:
    """源码运行时使用项目目录，打包后使用 exe 所在的可写目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def main() -> int:
    # 必须先于 Tk 和系统文件选择窗口创建，否则 Windows 会对整个界面做位图拉伸。
    enable_high_dpi_awareness()

    from tkinter import messagebox

    from src.config import AppPaths
    from src.gui import SalaryApp, TemplateSetupCancelled

    root = application_root()
    paths = AppPaths.from_root(root)
    try:
        app = SalaryApp(paths)
    except TemplateSetupCancelled:
        return 1
    except Exception as exc:
        messagebox.showerror("启动失败", str(exc))
        return 1
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
