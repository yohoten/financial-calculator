#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui.py —— 财务计算器图形界面（tkinter）
===========================================================================

在保留原命令行程序（main.py）的前提下，为其增加一个 tkinter 图形界面。

设计原则
--------
* **单一事实来源**：图形界面不重写计算与解读逻辑，而是调用 ``main.py`` 里已有的
  场景函数（scene_compound / scene_annuity / ... / run_all），并捕获其标准输出
  渲染到界面。因此命令行与界面的结果、口径、提示永远一致，不会出现"两套代码
  算出两个答案"。
* **界面与计算解耦**：所有耗时操作（计算 + Matplotlib 绘图 + Excel 导出）放在
  后台线程执行，主线程只负责刷新界面，避免窗口假死。
* **等宽文本对齐**：结果区使用中英文 2:1 等宽字体，保留命令行表格的对齐效果。

界面结构
--------
    总览页  一键运行全部场景 / 运行数值自检 / 打开输出目录 / 使用说明
    场景页  1 复利终值·现值   2 年金终值·现值   3 债权资本成本
            4 股权资本成本     5 WACC           6 贷款还款计划
    每个场景页：左侧参数表单 → 校验 → 计算 → 右侧结果文本 + 图表预览

运行方式
--------
    python main.py --gui            # 从命令行入口启动图形界面
    python gui.py                   # 直接启动图形界面
    双击 财务计算器GUI.pyw           # 无控制台窗口启动
    python gui.py --screenshot 路径  # 生成界面截图（用于报告插图）

命令行功能完全保留：``python main.py``（交互菜单）、``--demo``、``--scene N``。
"""

from __future__ import annotations

import contextlib
import io
import math
import os
import queue
import re
import sys
import threading
import time
import traceback
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import tkinter as tk
from tkinter import messagebox, ttk

import validators as vd

# 复用命令行主程序：场景函数、输出函数与输出目录全部与 CLI 同源
# ⚠️ 本模块的入口函数不得命名为 main，否则会遮蔽这个模块引用，
#    导致所有 main.OUT_DIR / main.scene_* 引用失败。入口统一叫 entry。
import main

try:                                                    # 可选：更高质量的图片缩放
    from PIL import Image, ImageTk
    HAS_PIL = True
except Exception:                                       # noqa: BLE001
    HAS_PIL = False


APP_TITLE = "财务计算器 · 资金时间价值与资本成本"
APP_VERSION = "v3.0 图形界面版"
APP_TAGLINE = "命令行与图形界面共用同一套计算、校验与解读内核"


# =============================================================================
# 一、外观
# =============================================================================
#: 深色工业仪表盘配色：降低大面积纯黑刺眼感，强调信息层级与数值对比度
UI = {
    "bg": "#0B1220", "panel": "#111C2F", "panel_alt": "#16253D",
    "header": "#101B30", "header_text": "#F8FAFC", "accent": "#F59E0B",
    "accent_soft": "#FDE68A", "h1": "#F8FAFC", "h2": "#7DD3FC",
    "ok": "#34D399", "warn": "#FBBF24", "err": "#FB7185",
    "muted": "#94A3B8", "body": "#E2E8F0", "border": "#263A58",
    "input": "#0F1A2E", "selection": "#1E3A5F",
}

#: 字体族在 Tk 根窗口创建后解析（见 resolve_fonts），此处为兜底值
FONTS = {"ui": "TkDefaultFont", "mono": "TkFixedFont", "num": "TkFixedFont"}


def resolve_fonts(root: tk.Misc) -> None:
    """
    探测本机可用字体。

    结果区必须使用"中文宽度 = 2 × 数字宽度"的字体（新宋体 / 宋体 / 黑体等），
    否则命令行输出的等宽表格在界面里会逐行错位。Consolas 等纯西文等宽字体
    的中文会回退到雅黑，宽度比约 1.78:1，不能用于表格排版。
    """
    from tkinter import font as tkfont
    try:
        families = set(tkfont.families(root))
    except Exception:                                   # noqa: BLE001
        return

    def pick(candidates: Sequence[str], fallback: str) -> str:
        for name in candidates:
            if name in families:
                return name
        return fallback

    FONTS["ui"] = pick(("微软雅黑", "Microsoft YaHei", "等线", "宋体"), "TkDefaultFont")
    FONTS["mono"] = pick(("新宋体", "NSimSun", "宋体", "SimSun", "Consolas"), "TkFixedFont")
    FONTS["num"] = pick(("Consolas", "Cascadia Mono", "Courier New"), "TkFixedFont")


def f_ui(size: int = 9, bold: bool = False):
    return (FONTS["ui"], size, "bold") if bold else (FONTS["ui"], size)


def f_mono(size: int = 12, bold: bool = False):
    return (FONTS["mono"], size, "bold") if bold else (FONTS["mono"], size)


def f_num(size: int = 18):
    return (FONTS["num"], size, "bold")


# =============================================================================
# 二、通用工具
# =============================================================================
_ANSI_RE = re.compile(r"\x1b\[[0-9;]*[A-Za-z]")


def strip_ansi(text: str) -> str:
    """去掉 ANSI 颜色码：命令行输出被界面接管后必须还原为纯文本。"""
    return _ANSI_RE.sub("", text or "")


def enable_dpi_awareness() -> None:
    """Windows 高分屏下避免界面模糊（非 Windows 平台直接跳过）。"""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:                                   # noqa: BLE001
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:                               # noqa: BLE001
            pass


def open_path(path: str) -> None:
    """用系统默认程序打开文件或目录。"""
    if not os.path.exists(path):
        messagebox.showwarning("文件不存在", f"找不到：\n{path}")
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)                          # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            import subprocess
            subprocess.Popen(["open", path])
        else:
            import subprocess
            subprocess.Popen(["xdg-open", path])
    except Exception as exc:                            # noqa: BLE001
        messagebox.showerror("无法打开", f"{path}\n\n{exc}")


# =============================================================================
# 三、执行适配层：把命令行场景函数变成可捕获的界面任务
# =============================================================================
@dataclass
class RunOutput:
    """一次界面任务的执行结果。"""
    name: str
    text: str = ""
    figures: List[str] = field(default_factory=list)
    files: List[str] = field(default_factory=list)
    summary: Any = None
    error: str = ""
    traceback_text: str = ""
    seconds: float = 0.0

    @property
    def ok(self) -> bool:
        return not self.error


def _collect_outputs(out_dir: str, since: float) -> Tuple[List[str], List[str]]:
    """收集 since 时刻之后被写入的输出文件，区分图片与其他产物。"""
    figures: List[str] = []
    files: List[str] = []
    try:
        names = sorted(os.listdir(out_dir))
    except OSError:
        return figures, files
    for name in names:
        path = os.path.join(out_dir, name)
        if not os.path.isfile(path):
            continue
        try:
            if os.path.getmtime(path) < since - 1.0:
                continue
        except OSError:
            continue
        (figures if name.lower().endswith(".png") else files).append(path)
    return figures, files


def capture(name: str, call: Callable[[], Any], out_dir: Optional[str] = None) -> RunOutput:
    """
    执行 call() 并捕获其标准输出与产物文件。

    call 内部就是 main.py 的场景函数（它们原样 print 到 stdout），
    这里把 stdout 重定向到内存，再交给界面渲染，因此界面与命令行结果完全一致。
    """
    target = out_dir or main.OUT_DIR
    started = time.time()
    buffer = io.StringIO()
    summary: Any = None
    error = ""
    trace = ""

    try:
        with contextlib.redirect_stdout(buffer):
            summary = call()
    except Exception as exc:                            # noqa: BLE001
        error = f"计算失败：{type(exc).__name__}: {exc}"
        trace = traceback.format_exc()
        summary = None

    figures, files = _collect_outputs(target, started)
    return RunOutput(name=name, text=strip_ansi(buffer.getvalue()), figures=figures,
                     files=files, summary=summary, error=error,
                     traceback_text=trace, seconds=time.time() - started)


# =============================================================================
# 四、结果文本渲染：按命令行输出的符号体系统一着色
# =============================================================================
_TITLE_RULE = set("═")
_TABLE_RULE = set("─━┈┄")


def classify_lines(text: str) -> List[Tuple[str, str]]:
    """
    把命令行输出转换成 (tag, line) 序列。

    识别依据是 main.py 约定的输出符号：
        ✔ / ⚠ / ✘  → 成功、警告、错误（行首符号优先判定）
        ▶          → 小节标题
        ═ 整行      → 一级标题框线，紧随其后的一行即标题
        ─ 整行      → 表格线（muted）
        4 空格缩进且下一行是表格线 → 表名（h2）

    注意：标题框线（═）与表格线（─）必须分开处理，否则表格首行会被误判为标题。
    """
    raw_lines = text.splitlines()
    out: List[Tuple[str, str]] = []
    expect_h1 = False
    just_h1 = False

    for i, raw in enumerate(raw_lines):
        s = raw.strip()
        nxt = raw_lines[i + 1].strip() if i + 1 < len(raw_lines) else ""

        if not s:
            out.append(("body", ""))
            continue

        # 1. 行首符号优先（不受前后文影响）
        if s.startswith("✔"):
            out.append(("ok", s))
            continue
        if s.startswith("⚠"):
            out.append(("warn", s))
            continue
        if s.startswith("✘"):
            out.append(("err", s))
            continue
        if s.startswith("▶"):
            out.append(("h2", s))
            continue

        # 2. 一级标题框线：整行都是 ═
        #    title() 用上下两条框线夹住标题：上线开启“期待标题”，
        #    下线不得再次开启，否则标题后的第一个非缩进行会被误判为标题。
        if len(s) >= 6 and set(s) <= _TITLE_RULE:
            out.append(("rule", raw))
            expect_h1 = not just_h1
            just_h1 = False
            continue
        if expect_h1 and not raw.startswith("    "):
            out.append(("h1", s))
            expect_h1 = False
            just_h1 = True
            continue

        # 3. 表格线：整行都是 ─（缩进不算）
        if len(s) >= 6 and set(s) <= (_TABLE_RULE | {" "}):
            out.append(("muted", raw))
            continue

        # 4. 表名：4 空格缩进 + 不含数字 + 下一行是表格线
        #    （表格最后一行数据同样“下一行是表格线”，用“不含数字”把它排除）
        if (raw.startswith("    ") and len(s) < 40 and nxt
                and not any(ch.isdigit() for ch in s)
                and len(nxt) >= 6 and set(nxt) <= (_TABLE_RULE | {" "})):
            out.append(("h2", s))
            continue

        out.append(("body", raw))

    return out


# =============================================================================
# 五、通用组件
# =============================================================================
class ResultView(ttk.Frame):
    """结果文本区：只读、可滚动、按语义着色。"""

    def __init__(self, master, height: int = 10):
        super().__init__(master)
        # 用 Text + 双向滚动条代替 ScrolledText：wrap="none" 保证命令行表格不被折行，
        # 超长的敏感性分析表可以横向滚动查看。
        # height 只是“最小行数”（实际高度由布局撑开），取小值可避免把窗口最小高度顶大，
        # 否则窗口放不下时底部状态栏会被压成 1px。
        self.text = tk.Text(
            self, wrap="none", height=height, font=f_mono(12),
            bg=UI["input"], fg=UI["body"], relief="solid", bd=1,
            insertbackground=UI["accent"], selectbackground=UI["selection"],
            padx=12, pady=10, highlightbackground=UI["border"], highlightthickness=1)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.text.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.text.configure(state="disabled")

        self.text.tag_configure("h1", font=f_ui(13, True), foreground=UI["accent"],
                                spacing1=8, spacing3=6)
        self.text.tag_configure("h2", font=f_ui(11, True), foreground=UI["h2"],
                                spacing1=8, spacing3=4)
        self.text.tag_configure("rule", font=f_mono(12), foreground=UI["border"])
        self.text.tag_configure("body", font=f_mono(12), foreground=UI["body"])
        self.text.tag_configure("ok", font=f_mono(12, True), foreground=UI["ok"])
        self.text.tag_configure("warn", font=f_mono(12), foreground=UI["warn"])
        self.text.tag_configure("err", font=f_mono(12, True), foreground=UI["err"])
        self.text.tag_configure("muted", font=f_mono(11), foreground=UI["muted"])

    def clear(self) -> None:
        self._write([])

    def _write(self, chunks: Sequence[Tuple[str, str]]) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        for tag, line in chunks:
            self.text.insert("end", line + "\n", tag)
        self.text.configure(state="disabled")
        self.text.see("1.0")

    def show_placeholder(self, title: str, lines: Sequence[Tuple[str, str]]) -> None:
        chunks: List[Tuple[str, str]] = [("h1", title), ("h1", "")]
        chunks += list(lines)
        self._write(chunks)

    def show_output(self, result: RunOutput) -> None:
        """
        渲染一次任务的输出。

        不再额外加一行标题：命令行输出本身就以 ═ 框线标题开头，
        重复添加会出现两个一模一样的标题。
        """
        chunks: List[Tuple[str, str]] = []
        if result.error:
            chunks += [("h1", result.name), ("err", result.error)]
            if result.traceback_text:
                chunks += [("body", ""), ("muted", result.traceback_text)]

        body = classify_lines(result.text)
        if body:
            chunks += body
        elif not result.error:
            chunks += [("muted", f"{result.name}：本次未产生文本输出。")]

        chunks += [("body", ""),
                   ("muted", f"耗时 {result.seconds:.2f} 秒 · 输出目录 {main.OUT_DIR}")]
        if result.figures or result.files:
            chunks += [("h2", "本次生成的产物")]
            chunks += [("ok", "  ✔ " + os.path.basename(p))
                       for p in result.figures + result.files]
        self._write(chunks)


class FigurePreview(ttk.LabelFrame):
    """图表预览：下拉选择输出目录中的 PNG，等比缩放显示。"""

    def __init__(self, master, title: str = "图表预览"):
        super().__init__(master, text=title, padding=8)

        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text="图表：").pack(side="left")
        self.combo = ttk.Combobox(top, state="readonly", width=32)
        self.combo.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.combo.bind("<<ComboboxSelected>>", lambda _e: self._on_pick())
        ttk.Button(top, text="刷新", width=6, command=self.refresh).pack(side="left", padx=2)
        ttk.Button(top, text="打开原图", width=9, command=self._open_current).pack(side="left", padx=2)

        self.image_label = tk.Label(self, bg=UI["input"], relief="solid", bd=1,
                                    highlightbackground=UI["border"], highlightthickness=1,
                                    text="计算后自动显示生成的图表", fg=UI["muted"],
                                    font=f_ui(10))
        self.image_label.pack(fill="both", expand=True, pady=(8, 0))

        self.hint = ttk.Label(self, text="—", foreground=UI["muted"])
        self.hint.pack(anchor="w", pady=(4, 0))

        self._photo: Optional[Any] = None               # 必须持有引用，否则被回收
        self._path: Optional[str] = None

    # ---------------- 内部 ----------------
    @staticmethod
    def _pngs() -> List[str]:
        try:
            names = [n for n in sorted(os.listdir(main.OUT_DIR)) if n.lower().endswith(".png")]
        except OSError:
            return []
        return [os.path.join(main.OUT_DIR, n) for n in names]

    def _on_pick(self) -> None:
        paths = self._pngs()
        idx = self.combo.current()
        if 0 <= idx < len(paths):
            self.show(paths[idx])

    def _open_current(self) -> None:
        open_path(self._path or main.OUT_DIR)

    # ---------------- 对外 ----------------
    def refresh(self, select: Optional[str] = None) -> None:
        """重新扫描输出目录；select 为文件名时优先选中它。"""
        paths = self._pngs()
        self.combo["values"] = [os.path.basename(p) for p in paths]
        target = None
        if select:
            for i, p in enumerate(paths):
                if os.path.basename(p) == select:
                    target = i
                    break
        if target is None and paths:
            target = len(paths) - 1
        if target is not None:
            self.combo.current(target)
            self.show(paths[target])
        elif not paths:
            self.image_label.configure(image="", text="输出目录中还没有图表",
                                       fg=UI["muted"], width=0, height=0)
            self.hint.configure(text="—")

    def show(self, path: str) -> None:
        self.image_label.update_idletasks()
        avail_w = max(self.image_label.winfo_width(), 320)
        avail_h = max(self.image_label.winfo_height(), 240)
        try:
            if HAS_PIL:
                img = Image.open(path)
                img.load()
                w, h = img.size
                scale = min(avail_w / w, avail_h / h, 1.0)
                if scale < 1.0:
                    img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))),
                                     Image.LANCZOS)
                self._photo = ImageTk.PhotoImage(img)
                factor_info = f"缩放 {scale:.0%}"
            else:
                photo = tk.PhotoImage(file=path)
                factor = max(1, math.ceil(max(photo.width() / avail_w,
                                              photo.height() / avail_h)))
                if factor > 1:
                    photo = photo.subsample(factor, factor)
                self._photo = photo
                factor_info = f"缩放 1/{factor}"
        except Exception as exc:                        # noqa: BLE001
            self.image_label.configure(image="", text=f"无法预览：{exc}", fg=UI["err"])
            self.hint.configure(text=os.path.basename(path))
            return

        self.image_label.configure(image=self._photo, text="")
        self._path = path
        try:
            size_kb = os.path.getsize(path) / 1024
        except OSError:
            size_kb = 0
        self.hint.configure(text=f"{os.path.basename(path)}   原图 {size_kb:.0f} KB · {factor_info}",
                            foreground=UI["muted"])


@dataclass
class FieldSpec:
    """参数输入字段定义（rule 对应 validators.RULES 中的规则名）。"""
    key: str
    label: str
    default: Any
    rule: Optional[str] = None
    kind: str = "number"                    # number | int | choice | text
    options: Sequence[Tuple[str, Any]] = ()
    hint: str = ""


# =============================================================================
# 六、场景页
# =============================================================================
class SceneTab(ttk.Frame):
    """一个场景页：左侧参数表单，中间结果文本，右侧图表预览。"""

    def __init__(self, master, app: "FinanceCalculatorApp", spec: Dict[str, Any]):
        super().__init__(master, padding=10)
        self.app = app
        self.spec = spec
        self.fields: List[FieldSpec] = list(spec["fields"])
        self.vars: Dict[str, Any] = {}
        self.entries: Dict[str, Any] = {}

        ttk.Label(self, text=spec["title"], font=f_ui(14, True),
                  foreground=UI["accent"]).pack(anchor="w")
        if spec.get("desc"):
            ttk.Label(self, text=spec["desc"], foreground=UI["muted"]).pack(anchor="w", pady=(2, 8))

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        middle = ttk.Frame(body)
        body.add(left, weight=4)
        body.add(middle, weight=5)

        self._build_form(left)

        self.result = ResultView(middle)
        self.result.pack(fill="both", expand=True)

        self.preview = FigurePreview(body, "图表预览")
        body.add(self.preview, weight=5)

    # ---------------- 参数表单 ----------------
    def _build_form(self, parent: ttk.Frame) -> None:
        """参数表单：一行一个字段（标签 | 输入框 | 口径提示），紧凑到一屏内。"""
        ttk.Label(parent, text="利率 / 税率填小数，期数 / 年数填正整数。",
                  foreground=UI["muted"], font=f_ui(8)).pack(anchor="w", pady=(0, 4))

        box = ttk.LabelFrame(parent, text="参数输入", padding=10)
        box.pack(fill="x")
        box.columnconfigure(1, weight=1)

        for i, spec in enumerate(self.fields):
            ttk.Label(box, text=spec.label).grid(row=i, column=0, sticky="w",
                                                 padx=(0, 10), pady=5)
            if spec.kind == "choice":
                var = tk.StringVar(value=self._label_of(spec, spec.default))
                widget = ttk.Combobox(box, textvariable=var, state="readonly", width=20,
                                      values=[lbl for lbl, _ in spec.options])
            else:
                var = tk.StringVar(value=self._fmt_default(spec.default))
                widget = ttk.Entry(box, textvariable=var, width=22)
            widget.grid(row=i, column=1, sticky="ew", pady=5)
            self.vars[spec.key] = var
            self.entries[spec.key] = widget
            if spec.hint:
                ttk.Label(box, text=spec.hint, foreground=UI["muted"],
                          font=f_ui(8)).grid(row=i, column=2, sticky="w", padx=(10, 0))

        btns = ttk.Frame(parent)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="计算并生成图表", style="Accent.TButton",
                   command=self.on_calc).pack(side="left")
        ttk.Button(btns, text="恢复示例值", command=self.on_reset).pack(side="left", padx=8)
        ttk.Button(btns, text="清空", command=self.on_clear).pack(side="left")

    @staticmethod
    def _fmt_default(value: Any) -> str:
        if isinstance(value, float):
            return f"{value:g}"
        return str(value)

    @staticmethod
    def _label_of(spec: FieldSpec, value: Any) -> str:
        for lbl, val in spec.options:
            if val == value:
                return lbl
        return spec.options[0][0] if spec.options else ""

    # ---------------- 读取与校验 ----------------
    def read_params(self) -> Tuple[Optional[Dict[str, Any]], List[str], List[str]]:
        """读取表单：返回 (参数, 错误, 提示)。错误非空时参数为 None。"""
        values: Dict[str, Any] = {}
        errors: List[str] = []
        warnings: List[str] = []

        for spec in self.fields:
            widget = self.entries[spec.key]
            if spec.kind != "choice":
                widget.configure(style="TEntry")

            raw = str(self.vars[spec.key].get()).strip()

            if spec.kind == "choice":
                mapping = {lbl: val for lbl, val in spec.options}
                values[spec.key] = mapping.get(raw, spec.default)
                continue

            if spec.kind == "text":
                # 文本型字段（如逗号分隔的现金流序列）：原样传参，
                # 由 main.py 的解析函数统一处理，错误也走同一条中文提示通道。
                values[spec.key] = raw
                continue

            if raw == "":
                errors.append(f"【{spec.label}】不能为空，请输入一个数值。")
                widget.configure(style="Error.TEntry")
                continue
            if spec.rule:
                res = vd.validate(spec.rule, raw)
                if res.errors:
                    errors += res.errors
                    widget.configure(style="Error.TEntry")
                    continue
                warnings += res.warnings
            try:
                num = float(raw)
            except ValueError:
                errors.append(f"【{spec.label}】必须为数字，当前输入为“{raw}”。")
                widget.configure(style="Error.TEntry")
                continue
            values[spec.key] = int(round(num)) if spec.kind == "int" else num

        return (values if not errors else None), errors, warnings

    # ---------------- 事件 ----------------
    def on_calc(self) -> None:
        values, errors, warnings = self.read_params()
        if values is None:
            chunks = [("err", e) for e in errors]
            chunks += [("body", ""),
                       ("muted", "提示：利率请填小数（6% 填 0.06）；期数 / 年数填正整数；金额填数值（元）。")]
            self.result.show_placeholder("参数校验未通过", chunks)
            self.app.set_status("参数校验未通过，已标红错误字段")
            return

        if warnings:
            self.result.show_placeholder("参数提示（不影响计算）",
                                         [("warn", w) for w in warnings])
        runner = self.spec["runner"]
        self.app.submit(self.spec["title"],
                        lambda params=values: runner(False, params), self)

    def on_reset(self) -> None:
        for spec in self.fields:
            if spec.kind == "choice":
                self.vars[spec.key].set(self._label_of(spec, spec.default))
            else:
                self.vars[spec.key].set(self._fmt_default(spec.default))
                self.entries[spec.key].configure(style="TEntry")

    def on_clear(self) -> None:
        for spec in self.fields:
            if spec.kind != "choice":
                self.vars[spec.key].set("")

    # ---------------- 结果回调 ----------------
    def show_result(self, result: RunOutput) -> None:
        self.result.show_output(result)
        if result.figures:
            self.preview.refresh(os.path.basename(result.figures[0]))
        else:
            self.preview.refresh()


# =============================================================================
# 七、总览页
# =============================================================================
class OverviewTab(ttk.Frame):
    """总览页：一键运行全部场景、数值自检、导出用例簿、打开输出目录。"""

    def __init__(self, master, app: "FinanceCalculatorApp"):
        super().__init__(master, padding=10)
        self.app = app

        hero = tk.Frame(self, bg=UI["panel"], padx=18, pady=14,
                        highlightbackground=UI["border"], highlightthickness=1)
        hero.pack(fill="x", pady=(0, 10))
        tk.Label(hero, text=APP_TITLE, bg=UI["panel"], fg=UI["accent"],
                 font=f_ui(19, True)).pack(anchor="w")
        tk.Label(hero, text=APP_TAGLINE, bg=UI["panel"], fg=UI["body"],
                 font=f_ui(10)).pack(anchor="w", pady=(4, 0))
        tk.Label(hero, text="第 3 章 货币时间价值   /   第 4 章 资本成本   /   可审计的图表与表格输出",
                 bg=UI["panel"], fg=UI["muted"], font=f_ui(9)).pack(anchor="w", pady=(8, 0))

        cards = tk.Frame(self, bg=UI["bg"])
        cards.pack(fill="x", pady=(0, 10))
        for label, value, color in (("计算场景", "06", UI["accent"]),
                                    ("专业图表", "08", UI["h2"]),
                                    ("校验层级", "L1·L2·L3", UI["ok"]),
                                    ("导出格式", "XLSX / MD", UI["accent_soft"])):
            self._stat_card(cards, label, value, color)

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="一键运行全部场景", style="Accent.TButton",
                   command=self.on_demo).pack(side="left")
        ttk.Button(bar, text="运行数值自检", command=self.on_selftest).pack(side="left", padx=8)
        ttk.Button(bar, text="打开输出目录",
                   command=lambda: open_path(main.OUT_DIR)).pack(side="left")
        ttk.Button(bar, text="使用说明", command=self.app.show_help).pack(side="left", padx=8)

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True)
        self.result = ResultView(body)
        body.add(self.result, weight=5)
        self.preview = FigurePreview(body, "产物预览")
        body.add(self.preview, weight=5)

        self.result.show_placeholder("欢迎使用", [
            ("body", f"版本 {APP_VERSION}，对应课程第 3 章与第 4 章。"),
            ("body", ""),
            ("body", "· 上方场景页：左侧填参数 → 「计算并生成图表」，右侧查看结果、解读与图表。"),
            ("body", "· 「一键运行全部场景」按示例参数跑完 6 个场景，生成 8 张图表、"),
            ("body", "  09_计算结果汇总.xlsx 与 计算结果汇总.md。"),
            ("body", "· 「运行数值自检」执行 7 组用例，逐项与教材系数表、解析解比对，"),
            ("body", "  同时导出 10_测试用例明细.xlsx。"),
            ("body", ""),
            ("muted", f"输出目录：{main.OUT_DIR}"),
            ("muted", "界面与命令行（python main.py）共用同一套计算内核，结果完全一致。"),
        ])

    @staticmethod
    def _stat_card(parent: tk.Frame, label: str, value: str, accent: str) -> None:
        card = tk.Frame(parent, bg=UI["panel_alt"], padx=14, pady=10,
                        highlightbackground=UI["border"], highlightthickness=1)
        card.pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Label(card, text=label, bg=UI["panel_alt"], fg=UI["muted"],
                 font=f_ui(9)).pack(anchor="w")
        tk.Label(card, text=value, bg=UI["panel_alt"], fg=accent,
                 font=f_num(17)).pack(anchor="w", pady=(5, 0))

    # ---------------- 事件 ----------------
    def on_demo(self) -> None:
        self.app.submit("一键运行全部场景", lambda: main.run_all(False), self)

    def on_selftest(self) -> None:
        self.app.submit("数值自检（并导出测试用例簿）", self._selftest_call, self)

    @staticmethod
    def _selftest_call() -> Any:
        """复用 test_cases.py 的用例集，避免维护第二套测试。"""
        import test_cases
        test_cases.main()
        return None

    def show_result(self, result: RunOutput) -> None:
        self.result.show_output(result)
        self.preview.refresh()
        if result.ok:
            messagebox.showinfo("运行完成",
                                f"{result.name} 已执行完毕。\n\n"
                                f"产物目录：\n{main.OUT_DIR}")


# =============================================================================
# 八、主窗口
# =============================================================================
class FinanceCalculatorApp:
    """主应用：标题条 + 工具条 + 功能页 + 状态栏；耗时任务在后台线程执行。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self._queue: "queue.Queue[Tuple[RunOutput, Any]]" = queue.Queue()
        self._busy = False

        resolve_fonts(root)

        root.title(f"{APP_TITLE} {APP_VERSION}")
        self._place_window()
        root.configure(bg=UI["bg"])
        root.report_callback_exception = self._on_tk_error

        self._setup_style()
        self._configure_window()
        self._build_header()
        self._build_toolbar()
        # 状态栏必须先于 Notebook 布局（side="bottom"），
        # 否则空间不足时它会被 expand 的 Notebook 挤成 1px 高、看不见。
        self._build_statusbar()
        self._build_notebook()

    # ---------------- 窗口与样式 ----------------
    def _place_window(self) -> None:
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        width = min(1440, max(1024, screen_w - 80))
        height = min(920, max(680, screen_h - 120))
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2 - 20)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(1024, 680)

    def _configure_window(self) -> None:
        self.root.option_add("*TCombobox*Listbox.background", UI["input"])
        self.root.option_add("*TCombobox*Listbox.foreground", UI["body"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", UI["selection"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", UI["body"])

    def _setup_style(self) -> None:
        """
        优先选 clam 主题：它对 background / foreground 支持最完整，能真正呈现
        深色仪表盘；vista / xpnative 更原生但会忽略大量 style 配置，仅作兜底。
        """
        style = ttk.Style()
        for theme in ("clam", "alt", "default", "vista", "xpnative"):
            if theme in style.theme_names():
                try:
                    style.theme_use(theme)
                    break
                except tk.TclError:
                    continue

        style.configure(".", font=f_ui(9), background=UI["bg"], foreground=UI["body"])
        for cls in ("TFrame", "TLabel", "TLabelframe"):
            style.configure(cls, background=UI["bg"], foreground=UI["body"])
        style.configure("TLabelframe", bordercolor=UI["border"], relief="solid")
        style.configure("TLabelframe.Label", background=UI["bg"], foreground=UI["muted"],
                        font=f_ui(9, True))

        style.configure("TButton", font=f_ui(9, True), padding=(12, 7),
                        background=UI["panel_alt"], foreground=UI["body"],
                        bordercolor=UI["border"])
        style.map("TButton",
                  background=[("active", UI["selection"]), ("pressed", UI["selection"]),
                              ("disabled", UI["panel"])],
                  foreground=[("disabled", UI["muted"]), ("active", UI["accent_soft"])])
        style.configure("Accent.TButton", font=f_ui(9, True), padding=(14, 8),
                        background=UI["accent"], foreground="#111827",
                        bordercolor=UI["accent"])
        style.map("Accent.TButton",
                  background=[("active", "#FBBF24"), ("pressed", "#D97706"),
                              ("disabled", UI["panel_alt"])],
                  foreground=[("disabled", UI["muted"])])

        style.configure("TEntry", fieldbackground=UI["input"], foreground=UI["body"],
                        insertcolor=UI["accent"], bordercolor=UI["border"])
        style.configure("Error.TEntry", fieldbackground="#4C1D2A", foreground="#FFE4E6",
                        bordercolor=UI["err"])
        style.configure("TCombobox", fieldbackground=UI["input"], foreground=UI["body"],
                        selectbackground=UI["selection"], selectforeground=UI["body"],
                        bordercolor=UI["border"])
        style.configure("TNotebook", background=UI["bg"], bordercolor=UI["border"],
                        tabmargins=(2, 6, 2, 0))
        style.configure("TNotebook.Tab", font=f_ui(9, True), padding=(14, 8),
                        background=UI["panel"], foreground=UI["muted"],
                        bordercolor=UI["border"])
        style.map("TNotebook.Tab",
                  background=[("selected", UI["accent"]), ("active", UI["panel_alt"])],
                  foreground=[("selected", "#111827"), ("active", UI["body"])])
        style.configure("TProgressbar", background=UI["accent"],
                        troughcolor=UI["panel"], bordercolor=UI["border"])

    def _build_header(self) -> None:
        head = tk.Frame(self.root, bg=UI["header"], height=74,
                        highlightbackground=UI["border"], highlightthickness=1)
        head.pack(fill="x")
        head.pack_propagate(False)

        left = tk.Frame(head, bg=UI["header"])
        left.pack(side="left", padx=18, pady=10)
        tk.Label(left, text="FINANCE / LAB", bg=UI["header"], fg=UI["accent"],
                 font=f_num(9)).pack(anchor="w")
        tk.Label(left, text=APP_TITLE, bg=UI["header"], fg=UI["header_text"],
                 font=f_ui(15, True)).pack(anchor="w", pady=(2, 0))
        tk.Label(head, text="TVM  ·  CAPM  ·  WACC  ·  LOAN", bg=UI["header"],
                 fg=UI["muted"], font=f_num(9)).pack(side="left", padx=22)
        tk.Label(head, text=APP_VERSION, bg=UI["header"], fg=UI["accent_soft"],
                 font=f_num(9)).pack(side="right", padx=18)

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self.root, padding=(14, 10, 14, 7))
        bar.pack(fill="x")
        ttk.Button(bar, text="一键运行全部场景", style="Accent.TButton",
                   command=self.run_demo).pack(side="left")
        ttk.Button(bar, text="数值自检", command=self.run_selftest).pack(side="left", padx=8)
        ttk.Button(bar, text="打开输出目录",
                   command=lambda: open_path(main.OUT_DIR)).pack(side="left")
        ttk.Button(bar, text="使用说明", command=self.show_help).pack(side="left", padx=8)
        ttk.Label(bar, text="输入 → 校验 → 计算 → 解释 → 导出",
                  foreground=UI["muted"], font=f_ui(9)).pack(side="right")

    def _build_notebook(self) -> None:
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self.overview = OverviewTab(self.nb, self)
        self.nb.add(self.overview, text="总览")

        self.tabs: Dict[str, SceneTab] = {}
        for spec in SCENE_SPECS:
            tab = SceneTab(self.nb, self, spec)
            self.nb.add(tab, text=spec["tab"])
            self.tabs[spec["tab"]] = tab

    def _build_statusbar(self) -> None:
        bar = ttk.Frame(self.root, padding=(10, 2, 10, 6))
        bar.pack(fill="x", side="bottom")
        self.status = ttk.Label(bar, text="就绪", foreground=UI["muted"])
        self.status.pack(side="left")
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=140)
        self.progress.pack(side="left", padx=12)
        ttk.Label(bar, text=f"输出目录：{main.OUT_DIR}",
                  foreground=UI["muted"]).pack(side="right")

    # ---------------- 对外动作 ----------------
    def set_status(self, text: str) -> None:
        self.status.configure(text=text)

    def run_demo(self) -> None:
        self.nb.select(self.overview)
        self.overview.on_demo()

    def run_selftest(self) -> None:
        self.nb.select(self.overview)
        self.overview.on_selftest()

    def submit(self, name: str, call: Callable[[], Any], view: Any) -> None:
        """把耗时任务放到后台线程执行，完成后回到主线程刷新界面。"""
        if self._busy:
            messagebox.showinfo("正在执行", "上一个任务尚未完成，请稍候。")
            return

        self._busy = True
        self.set_status(f"正在执行：{name} …")
        self.progress.start(90)

        def worker() -> None:
            result = capture(name, call)
            self._queue.put((result, view))

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(80, self._poll_queue)

    def _poll_queue(self) -> None:
        try:
            result, view = self._queue.get_nowait()
        except queue.Empty:
            self.root.after(80, self._poll_queue)
            return

        self.progress.stop()
        self._busy = False
        # 状态栏统一在这里收尾：任务完成的提示不能只写在各个页面的回调里，
        # 否则场景页算完后状态栏会一直停在“正在执行…”。
        if result.ok:
            self.set_status(f"{result.name} 完成 · 耗时 {result.seconds:.2f} 秒")
        else:
            self.set_status(f"{result.name} 失败：{result.error}")
        try:
            view.show_result(result)
        except Exception:                               # noqa: BLE001
            messagebox.showerror("界面刷新异常", traceback.format_exc()[-1500:])

    def show_help(self) -> None:
        messagebox.showinfo("使用说明", "\n".join([
            f"{APP_TITLE} {APP_VERSION}",
            "",
            "1. 六个场景页：左侧填参数 → 「计算并生成图表」，",
            "   中间显示计算结果与业务解读，右侧预览本次生成的图表。",
            "2. 参数口径：利率 / 税率 / 增长率填小数（6% → 0.06）；",
            "   期数 / 年数填正整数；金额默认单位为元；",
            "   WACC 页的资本结构金额按万元填报（保持同一口径即可）。",
            "3. 输入错误会被三层校验拦截（类型 → 边界 → 合理性），",
            "   错误框标红并给出“错在哪里 + 为什么错 + 应如何修正”的提示。",
            "4. 总览页可一键跑完全部示例场景，生成 8 张图表与 Excel / Markdown 汇总；",
            "   也可运行数值自检并导出测试用例簿。",
            "",
            "命令行功能保持不变：",
            "   python main.py              交互式菜单",
            "   python main.py --demo       一键运行全部场景",
            "   python main.py --scene 3    直接运行场景 3",
            "",
            f"输出目录：{main.OUT_DIR}",
            "界面与命令行共用同一计算内核，同一输入的输出一致。",
        ]))

    def _on_tk_error(self, exc, val, tb) -> None:
        text = "".join(traceback.format_exception(exc, val, tb))
        try:
            messagebox.showerror("程序异常", text[-1800:])
        except Exception:                               # noqa: BLE001
            print(text, file=sys.stderr)


# =============================================================================
# 九、七个场景的表单定义（runner 直接复用 main.py 的场景函数）
# =============================================================================
SCENE_SPECS: List[Dict[str, Any]] = [
    dict(
        tab="1 复利终值·现值",
        title="场景 1 ｜ 复利终值 / 现值计算（第 3 章）",
        desc="复利终值 FV = PV × (1 + r/m)^(m·n)；同时给出实际年利率 EAR 与计息频率效应。",
        runner=main.scene_compound,
        fields=[
            FieldSpec("pv", "现值 PV / 终值（元）", 100000, "positive_amount"),
            FieldSpec("rate", "年利率 r", 0.06, "rate", hint="小数"),
            FieldSpec("years", "年数 n", 10, "years", kind="int", hint="正整数"),
            FieldSpec("freq", "每年计息次数 m", 1, None, kind="choice",
                      options=[("1（按年）", 1), ("2（按半年）", 2), ("4（按季）", 4),
                               ("12（按月）", 12), ("365（按日）", 365)]),
            FieldSpec("mode", "计算方向", "1", None, kind="choice",
                      options=[("已知现值求终值", "1"), ("已知终值求现值", "2")]),
        ],
    ),
    dict(
        tab="2 年金终值·现值",
        title="场景 2 ｜ 年金终值 / 现值计算（第 3 章）",
        desc="普通年金与预付年金对照，展示每期现金流的时间轴与折现贡献分解。",
        runner=main.scene_annuity,
        fields=[
            FieldSpec("pmt", "每期现金流 PMT（元）", 10000, "positive_amount"),
            FieldSpec("rate", "每期利率 r", 0.08, "rate", hint="小数"),
            FieldSpec("periods", "期数 n", 10, "periods", kind="int", hint="正整数"),
            FieldSpec("periods_per_year", "每年收付次数", 1, None, kind="choice",
                      options=[("1（按年）", 1), ("2（按半年）", 2),
                               ("4（按季）", 4), ("12（按月）", 12)]),
            FieldSpec("due", "收付方式", False, None, kind="choice",
                      options=[("期末收付（普通年金）", False), ("期初收付（预付年金）", True)]),
        ],
    ),
    dict(
        tab="3 债权资本成本",
        title="场景 3 ｜ 债权资本成本计算（第 4 章）",
        desc="IRR / YTM 精确法与教材近似公式双引擎互验，并量化税盾效应。",
        runner=main.scene_debt,
        fields=[
            FieldSpec("face", "债券面值 F（元）", 1000, "positive_amount"),
            FieldSpec("coupon", "票面利率（年）", 0.06, "coupon_rate", hint="小数"),
            FieldSpec("years", "债券期限（年）", 5, "years", kind="int", hint="正整数"),
            FieldSpec("price", "发行价（元）", 950, "positive_amount"),
            FieldSpec("tax", "所得税税率", 0.25, "tax_rate", hint="一般企业 0.25"),
            FieldSpec("pay_freq", "每年付息次数", 1, None, kind="choice",
                      options=[("1（按年付息）", 1), ("2（按半年付息）", 2), ("4（按季付息）", 4)]),
            FieldSpec("flotation", "筹资费率 f", 0.0, None, hint="无则填 0"),
        ],
    ),
    dict(
        tab="4 股权资本成本",
        title="场景 4 ｜ 股权资本成本计算（第 4 章）",
        desc="CAPM 为主算法，DDM（戈登增长模型）作交叉验证，两法差异直接给出结论。",
        runner=main.scene_equity,
        fields=[
            FieldSpec("rf", "无风险利率 Rf", 0.025, "rf", hint="小数"),
            FieldSpec("beta", "β 系数", 1.2, "beta", hint="市场组合 = 1.0"),
            FieldSpec("rm", "市场期望收益率 Rm", 0.095, "rm", hint="小数"),
            FieldSpec("size_premium", "规模 / 流动性溢价", 0.0, None, hint="无则填 0"),
            FieldSpec("use_ddm", "是否用 DDM 交叉验证", True, None, kind="choice",
                      options=[("是（推荐）", True), ("否", False)]),
            FieldSpec("d1", "下一年度每股股利 D1（元）", 0.5, "positive_amount"),
            FieldSpec("p0", "股票现价 P0（元）", 12.0, "positive_amount"),
            FieldSpec("g", "股利增长率 g", 0.04, "growth", hint="应小于 Ks"),
        ],
    ),
    dict(
        tab="5 WACC",
        title="场景 5 ｜ 加权平均资本成本 WACC（第 4 章）",
        desc="债务成本 + 股权成本 + 加权，附资本结构优化模拟与敏感性热力图。",
        runner=main.scene_wacc,
        fields=[
            FieldSpec("face", "债券面值（元）", 1000, "positive_amount"),
            FieldSpec("coupon", "票面年利率", 0.06, "coupon_rate", hint="小数"),
            FieldSpec("years", "期限（年）", 5, "years", kind="int", hint="正整数"),
            FieldSpec("price", "发行价（元）", 980, "positive_amount"),
            FieldSpec("tax", "所得税税率", 0.25, "tax_rate", hint="小数"),
            FieldSpec("pay_freq", "每年付息次数", 1, None, kind="choice",
                      options=[("1（按年付息）", 1), ("2（按半年付息）", 2)]),
            FieldSpec("flotation", "筹资费率 f", 0.0, None, hint="无则填 0"),
            FieldSpec("rf", "无风险利率 Rf", 0.025, "rf", hint="小数"),
            FieldSpec("beta", "β 系数", 1.15, "beta"),
            FieldSpec("rm", "市场期望收益率 Rm", 0.095, "rm", hint="小数"),
            FieldSpec("equity", "股权市场价值 E（万元）", 6000, "positive_amount"),
            FieldSpec("debt", "债务市场价值 D（万元）", 4000, "positive_amount"),
            FieldSpec("pref", "优先股市场价值 P（万元）", 0.0, None, hint="无则 0"),
            FieldSpec("kp", "优先股资本成本 Kp", 0.0, None, hint="无则 0"),
        ],
    ),
    dict(
        tab="6 贷款还款计划",
        title="场景 6 ｜ 贷款还款计划（第 3 章应用）",
        desc="等额本息是年金现值公式的逆向应用：本金 = 各期还款额的年金现值。",
        runner=main.scene_loan,
        fields=[
            FieldSpec("principal", "贷款本金（元）", 1000000, "positive_amount"),
            FieldSpec("rate", "年利率", 0.042, "rate", hint="小数"),
            FieldSpec("years", "贷款年限", 20, "years", kind="int", hint="正整数"),
            FieldSpec("freq", "每年还款次数", 12, None, kind="choice",
                      options=[("12（按月）", 12), ("4（按季）", 4),
                               ("2（按半年）", 2), ("1（按年）", 1)]),
        ],
    ),
    dict(
        tab="7 永续年金·现金流",
        title="场景 7 ｜ 永续年金现值与自定义现金流序列（第 3 章）",
        desc="永续年金 PVP = PMT/(r-g)；自定义现金流序列支持任意期数与任意金额，"
             "同时给出滚动终值与折现净现值（两法互验）。",
        runner=main.scene_cashflow,
        fields=[
            FieldSpec("pmt", "永续年金·每期金额（元）", 50000, "positive_amount"),
            FieldSpec("rate", "永续年金·折现率 r", 0.08, "rate", hint="须大于 g"),
            FieldSpec("growth", "永续年金·增长率 g", 0.02, "growth", hint="无增长填 0"),
            FieldSpec("cashflows", "自定义现金流序列",
                      "-100000,30000,35000,40000,45000", None, kind="text",
                      hint="逗号分隔；投入负数、收回正数"),
            FieldSpec("cf_rate", "现金流·每期折现率 r", 0.08, "rate", hint="与序列同口径"),
            FieldSpec("periods_per_year", "现金流·每年发生次数", 1, None, kind="choice",
                      options=[("1（按年）", 1), ("2（按半年）", 2),
                               ("4（按季）", 4), ("12（按月）", 12)]),
        ],
    ),
]


# =============================================================================
# 十、入口
# =============================================================================
def launch() -> int:
    """启动图形界面（阻塞至窗口关闭）。"""
    enable_dpi_awareness()
    root = tk.Tk()
    FinanceCalculatorApp(root)
    root.mainloop()
    return 0


def screenshot(path: str) -> int:
    """无人工干预地渲染主窗口并截图，用于报告插图或界面回归验证。"""
    if not HAS_PIL:
        print("截图功能需要 Pillow：pip install pillow", file=sys.stderr)
        return 2
    enable_dpi_awareness()
    root = tk.Tk()
    FinanceCalculatorApp(root)
    root.attributes("-topmost", True)
    root.lift()

    def _shot() -> None:
        root.update_idletasks()
        root.update()
        try:
            from PIL import ImageGrab
            x, y = root.winfo_rootx(), root.winfo_rooty()
            w, h = root.winfo_width(), root.winfo_height()
            img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
            os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
            img.save(path)
            print(f"界面截图已保存：{path}")
        except Exception as exc:                        # noqa: BLE001
            print(f"截图失败：{type(exc).__name__}: {exc}", file=sys.stderr)
        finally:
            root.destroy()

    root.after(1200, _shot)
    root.mainloop()
    return 0


def entry(argv: Optional[Sequence[str]] = None) -> int:
    """图形界面入口：默认启动窗口，支持 --screenshot / --check。"""
    argv = list(sys.argv[1:] if argv is None else argv)

    if "--screenshot" in argv:
        idx = argv.index("--screenshot")
        target = argv[idx + 1] if idx + 1 < len(argv) else os.path.join(
            main.OUT_DIR, "gui_预览.png")
        return screenshot(target)

    if "--check" in argv:                                # 依赖与字体自检
        root = tk.Tk()
        root.withdraw()
        resolve_fonts(root)
        print(f"tkinter    ✔ {tk.TkVersion}")
        print(f"Pillow     {'✔ 可用' if HAS_PIL else '✘ 未安装（图表缩放将退化）'}")
        print(f"界面字体    {FONTS['ui']}")
        print(f"等宽字体    {FONTS['mono']}")
        print(f"输出目录    {main.OUT_DIR}")
        root.destroy()
        return 0

    return launch()


if __name__ == "__main__":
    sys.exit(entry())
