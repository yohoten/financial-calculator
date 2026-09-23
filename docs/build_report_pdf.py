# -*- coding: utf-8 -*-
"""
build_report_pdf.py —— 把《开发报告.md》导出为《开发报告.pdf》
================================================================

用途
----
报告正文以 Markdown 维护，PDF 是交付附件。只要 Markdown 改了，就应重出 PDF，
否则会出现"PDF 与正文说法不一致"的交付事故。本脚本把该步骤固化，避免每次
重新摸索工具链。

依赖
----
1. pandoc（本机实测 3.10）
2. Chrome 或 Edge（任一 Chromium 内核浏览器，用于无头打印）
3. 中文字体（Microsoft YaHei / SimSun，Windows 自带）

用法
----
    python docs/build_report_pdf.py                    # 使用默认浏览器探测顺序
    python docs/build_report_pdf.py --browser edge      # 指定 Edge
    python docs/build_report_pdf.py --md 其他文档.md     # 指定源文件
    python docs/build_report_pdf.py --keep-html         # 保留中间 HTML 便于排查

实现说明（几个踩过的坑，改脚本前先看）
--------------------------------------
* **公式**：正文含大量 LaTeX 公式。这里让 pandoc 输出 **MathML**（`--mathml`），
  由 Chromium 原生渲染——不依赖 KaTeX/MathJax 的 CDN，**离线可用**。
  PDF 中可见 Cambria Math 字体即是证据。
* **图片**：Markdown 里的插图写成 `figures/xx.png`（相对 docs/）。
  中间 HTML 若放到临时目录，相对路径会失效，因此注入 `<base href="...">`
  指向 docs/，这样中间文件放哪都能正确取图。
* **路径**：项目路径含全角括号 `（8）`，交给 Chrome 时须用**正斜杠 + 盘符**
  （`F:/...`）形式；在 Git Bash 下不要依赖 shell 的路径自动转换。
* **落盘时序**：Chrome 写 PDF 是异步的。脚本内以"文件存在且 mtime 更新"轮询确认，
  不要在启动 Chrome 后立刻用 `ls` 判断成败（Windows 目录缓存会给出旧时间戳）。
"""

from __future__ import annotations

import argparse
import pathlib
import shutil
import subprocess
import sys
import tempfile
import time

HERE = pathlib.Path(__file__).resolve().parent           # .../docs
DEFAULT_MD = HERE / "开发报告.md"
DEFAULT_PDF = HERE / "开发报告.pdf"
CSS = HERE / "report_print.css"

BROWSERS = {
    "chrome": [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ],
    "edge": [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ],
}

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<base href="{base}/">
<title>{title}</title>
<style>
{css}
</style>
</head>
<body>
{body}
</body>
</html>
"""


def find_pandoc() -> str:
    exe = shutil.which("pandoc")
    if exe:
        return exe
    for cand in (
        pathlib.Path.home() / "AppData/Local/Pandoc/pandoc.exe",
        pathlib.Path(r"C:\Program Files\Pandoc\pandoc.exe"),
    ):
        if cand.is_file():
            return str(cand)
    raise SystemExit("未找到 pandoc。请先安装：winget install --id JohnMacFarlane.Pandoc")


def find_browser(prefer: str = "auto") -> str:
    order = [prefer] if prefer in BROWSERS else []
    order += [k for k in ("chrome", "edge") if k not in order]
    for name in order:
        for path in BROWSERS[name]:
            if pathlib.Path(path).is_file():
                return path
    raise SystemExit("未找到 Chrome / Edge。请安装任一 Chromium 内核浏览器，或用 --browser 指定。")


def build_html(md: pathlib.Path, pandoc: str, css: str, workdir: pathlib.Path) -> pathlib.Path:
    """Markdown --(pandoc)--> HTML 片段 --(套样式)--> 完整 HTML。"""
    fragment = workdir / "body.html"
    subprocess.run(
        [
            pandoc,
            str(md),
            "-f", "gfm+tex_math_dollars",
            "-t", "html5",
            "--mathml",            # 公式转 MathML，离线可渲染
            "--wrap=none",
            "-o", str(fragment),
        ],
        check=True,
    )
    title = "货币时间价值与资本成本自动化计算器 · 开发报告"
    # 用 as_uri() 得到浏览器可识别的 file:/// 形式，避免手写路径的编码问题
    full = workdir / "report.html"
    full.write_text(
        HTML_TEMPLATE.format(
            base=HERE.as_uri(),
            title=title,
            css=css,
            body=fragment.read_text(encoding="utf-8"),
        ),
        encoding="utf-8",
    )
    return full


def print_pdf(browser: str, html: pathlib.Path, pdf: pathlib.Path,
              workdir: pathlib.Path, timeout: float = 180.0) -> None:
    """调用 Chromium 无头打印。以 mtime 轮询确认落盘，规避目录缓存导致的误判。"""
    before = pdf.stat().st_mtime if pdf.exists() else 0.0
    if pdf.exists():
        pdf.unlink()

    subprocess.run(
        [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--allow-file-access-from-files",
            f"--user-data-dir={workdir / 'chrome-profile'}",
            "--no-pdf-header-footer",
            "--virtual-time-budget=30000",
            f"--print-to-pdf={pdf.as_posix()}",
            html.as_uri(),
        ],
        check=False,
        capture_output=True,
    )

    deadline = time.time() + timeout
    while time.time() < deadline:
        if pdf.exists() and pdf.stat().st_size > 0:
            try:
                with pdf.open("rb") as fh:
                    if fh.read(5) == b"%PDF-":
                        if pdf.stat().st_mtime > before:
                            return
            except OSError:
                pass
        time.sleep(0.5)
    raise SystemExit(
        f"未在 {timeout:.0f}s 内生成有效的 PDF：{pdf}\n"
        f"排查建议：确认 {browser} 未被安全软件拦截；或加 --keep-html 后手工用浏览器打印。"
    )


def report_stats(pdf: pathlib.Path) -> str:
    """粗读 PDF，给出页数与嵌入图片数的自检信息（不依赖第三方 PDF 库）。"""
    raw = pdf.read_bytes()
    pages = 0
    for line in raw.split(b"/Count "):
        digits = b""
        for ch in line:
            if 48 <= ch <= 57:
                digits += bytes([ch])
            else:
                break
        if digits:
            pages = max(pages, int(digits))
    images = raw.count(b"/Subtype /Image") + raw.count(b"/Subtype/Image")
    return f"页数约 {pages or '未知'}，嵌入图片 {images} 张，大小 {pdf.stat().st_size / 1024 / 1024:.2f} MB"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="把《开发报告.md》导出为 PDF")
    ap.add_argument("--md", default=str(DEFAULT_MD), help="源 Markdown 文件")
    ap.add_argument("--out", default=str(DEFAULT_PDF), help="输出 PDF 路径")
    ap.add_argument("--browser", default="auto", choices=["auto", "chrome", "edge"],
                    help="使用的浏览器内核")
    ap.add_argument("--keep-html", action="store_true", help="保留中间 HTML 与工作目录")
    args = ap.parse_args(argv)

    md = pathlib.Path(args.md)
    pdf = pathlib.Path(args.out)
    if not md.is_file():
        raise SystemExit(f"源文件不存在：{md}")
    if not CSS.is_file():
        raise SystemExit(f"缺少样式文件：{CSS}")

    pandoc = find_pandoc()
    browser = find_browser(args.browser)
    print(f"[1/3] pandoc  : {pandoc}")
    print(f"[2/3] browser : {browser}")

    workdir = pathlib.Path(tempfile.mkdtemp(prefix="report-pdf-"))
    try:
        html = build_html(md, pandoc, CSS.read_text(encoding="utf-8"), workdir)
        print("[3/3] 正在打印 PDF …")
        print_pdf(browser, html, pdf, workdir)
        print(f"完成：{pdf}")
        print(f"      自检：{report_stats(pdf)}")
        if args.keep_html:
            print(f"      中间文件保留在：{workdir}")
    finally:
        if not args.keep_html:
            shutil.rmtree(workdir, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
