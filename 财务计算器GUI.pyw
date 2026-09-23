# -*- coding: utf-8 -*-
"""
财务计算器GUI.pyw —— 图形界面双击启动器（pythonw，无控制台窗口）
================================================================

作用
----
把"用 pythonw 以图形方式启动 gui.py"这件事封装成一个可以双击的文件：

    pythonw 财务计算器GUI.pyw     ← 双击等价于此

为什么需要单独一个文件
----------------------
1. `.py` 默认由 python.exe 关联，双击会弹出黑色控制台窗口；`.pyw` 由 pythonw.exe
   关联，双击无控制台——这是 Windows 上唯一可靠的"双击无窗口"方式。
2. pythonw 下 `sys.stdout` / `sys.stderr` 为 `None`，任何 `print` 或
   `sys.stdout.isatty()` 都会直接抛异常。本启动器把标准流兜底为可丢弃的
   空设备，再把真实错误写入同目录的启动日志，避免"双击后毫无反应"。

命令行等价写法
--------------
    python main.py --gui
    python gui.py
    python gui.py --check            # 依赖与字体自检
    python gui.py --screenshot 路径   # 生成界面截图
"""

from __future__ import annotations

import os
import sys
import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)


def _ensure_std_streams() -> None:
    """pythonw 下 stdout/stderr 为 None：补一个只写的空设备，避免一导入就崩。"""
    for name in ("stdout", "stderr"):
        if getattr(sys, name, None) is None:
            try:
                setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
            except OSError:                                  # noqa: BLE001
                pass


def main() -> int:
    _ensure_std_streams()
    try:
        import gui
    except Exception:                                        # noqa: BLE001
        _write_log(traceback.format_exc())
        _notify("图形界面启动失败",
                "依赖缺失或代码错误，详情见：\n"
                f"{os.path.join(BASE_DIR, '启动错误.log')}\n\n"
                "可先在命令行执行 python gui.py --check 查看环境自检结果。")
        return 1

    try:
        return gui.entry([])
    except Exception:                                        # noqa: BLE001
        _write_log(traceback.format_exc())
        _notify("图形界面运行异常",
                f"详情见：{os.path.join(BASE_DIR, '启动错误.log')}")
        return 1


def _write_log(text: str) -> None:
    try:
        with open(os.path.join(BASE_DIR, "启动错误.log"), "a", encoding="utf-8") as fh:
            fh.write(text + "\n" + "-" * 60 + "\n")
    except OSError:
        pass


def _notify(title: str, message: str) -> None:
    """尽量用弹窗告知用户——pythonw 下没有控制台，不弹窗就等于"双击无反应"。"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()
    except Exception:                                        # noqa: BLE001
        pass


if __name__ == "__main__":
    sys.exit(main())
