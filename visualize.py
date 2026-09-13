# -*- coding: utf-8 -*-
"""
visualize.py —— Matplotlib 可视化模块
======================================

对应作业要求"（三）可视化输出"：
  1. 复利增长曲线：现值 → 终值的指数增长轨迹，叠加不同利率对照，展示复利效应
  2. 年金现金流时间轴：用 stem/箭头图直观呈现每期现金流方向与金额
  3. WACC 结果图：资本结构饼图 + 各来源成本贡献条形图（双联图）
  4. 敏感性分析热力图：股权成本 × 债权成本 → WACC，辅助融资决策

工程化细节
----------
* 中文字体自动探测（Noto Sans CJK / 微软雅黑 / 黑体），并处理负号显示
* 统一色板与 rcParams，保证输出风格一致，适合直接放进报告
* 所有函数签名返回文件路径，便于批处理汇总
"""

from __future__ import annotations

import os
from typing import Dict, List, Optional, Sequence, Tuple

import matplotlib
matplotlib.use("Agg")                     # 无 GUI 环境（服务器/CI）下安全出图
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.ticker import FuncFormatter

# -----------------------------------------------------------------------------
# 全局样式与中文字体
# -----------------------------------------------------------------------------
_CJK_CANDIDATES = [
    "Noto Sans CJK SC", "Noto Sans CJK JP", "Noto Serif CJK SC", "Noto Serif CJK JP",
    "Source Han Sans SC", "WenQuanYi Zen Hei", "Microsoft YaHei", "SimHei", "PingFang SC",
]


def _setup_style() -> str:
    """
    探测可用中文字体并统一 rcParams，返回实际使用的字体名。

    注意：Linux 上的 Noto CJK 字体族常只注册了 "Noto Sans CJK JP" 这一名称
    （即使包含全部简繁字形），因此候选列表按"可用性优先"排序，
    只要该字体覆盖中文即可正确显示简体汉字。
    """
    from matplotlib import font_manager
    installed = {f.name for f in font_manager.fontManager.ttflist}
    chosen = next((n for n in _CJK_CANDIDATES if n in installed), "DejaVu Sans")

    plt.rcParams.update({
        "font.sans-serif": [chosen, "DejaVu Sans"],
        "font.family": "sans-serif",
        "axes.unicode_minus": False,          # 解决负号显示为方块
        "figure.dpi": 130,
        "savefig.dpi": 160,
        "savefig.bbox": "tight",
        "axes.edgecolor": "#B8C2CC",
        "axes.labelcolor": "#2F3B52",
        "axes.titlesize": 14,
        "axes.titleweight": "bold",
        "text.color": "#2F3B52",
        "xtick.color": "#5A6B85",
        "ytick.color": "#5A6B85",
        "grid.color": "#E3E8EF",
        "grid.linewidth": 0.8,
        "legend.frameon": False,
        "figure.facecolor": "white",
        "axes.facecolor": "white",
    })
    return chosen


CJK_FONT = _setup_style()

# 统一色板
C = {
    "primary": "#2563EB",   # 主蓝
    "accent": "#F59E0B",    # 强调橙
    "green": "#10B981",     # 增长绿
    "red": "#EF4444",       # 风险红
    "purple": "#8B5CF6",    # 紫
    "gray": "#94A3B8",      # 辅助灰
    "dark": "#1E293B",
}
PALETTE = [C["primary"], C["accent"], C["green"], C["purple"], C["red"], C["gray"]]


def _ensure_dir(path: str) -> str:
    d = os.path.dirname(os.path.abspath(path))
    os.makedirs(d, exist_ok=True)
    return path


def _pct(x, _pos=None) -> str:
    return f"{x:.0%}"


def _money(x, _pos=None) -> str:
    """金额刻度：超过 1 万自动切换为“万”单位。"""
    if abs(x) >= 1e8:
        return f"{x / 1e8:.1f}亿"
    if abs(x) >= 1e4:
        return f"{x / 1e4:.0f}万"
    return f"{x:,.0f}"


# =============================================================================
# 图 1：复利增长曲线
# =============================================================================
def plot_compound_growth(
    pv: float,
    rate: float,
    years: int,
    out_path: str,
    compare_rates: Optional[Sequence[float]] = None,
    freq: int = 1,
    title: str = "复利增长曲线：货币的时间价值",
) -> str:
    """
    绘制复利增长曲线。

    左侧主图：本金 + 利息堆积面积图，直观展示"利息超过本金"的临界点（翻倍点）。
    右侧副图：不同利率下的终值对照条形图，展示利率对复利的放大作用。

    参数
    ----
    pv            : 现值（期初本金）
    rate          : 年化利率（小数）
    years         : 年数
    compare_rates : 对照利率列表，如 [0.02, 0.04, 0.06]，默认自动生成
    freq          : 每年计息次数
    """
    _ensure_dir(out_path)
    t = np.arange(0, years + 1)
    factor = (1.0 + rate / freq) ** (freq * t)
    values = pv * factor

    fig = plt.figure(figsize=(13, 5.4))
    gs = fig.add_gridspec(1, 2, width_ratios=[2.05, 1.0], wspace=0.24)

    # ---------- 左：堆积面积图 ----------
    ax = fig.add_subplot(gs[0, 0])
    ax.fill_between(t, 0, pv, color=C["primary"], alpha=0.85, label=f"本金 {pv:,.0f} 元")
    ax.fill_between(t, pv, values, color=C["green"], alpha=0.75, label="累计利息收益")
    ax.plot(t, values, color=C["dark"], linewidth=2.0, zorder=5, label="本息合计")

    # 标注翻倍点（首次达到本金 2 倍的时点）
    double_idx = np.where(values >= 2 * pv)[0]
    if len(double_idx) > 0:
        k = double_idx[0]
        ax.scatter([t[k]], [values[k]], color=C["red"], s=62, zorder=6)
        ax.annotate(
            f"资金翻倍\n第 {t[k]} 年 · {values[k]:,.0f} 元",
            xy=(t[k], values[k]), xytext=(t[k] + max(years * 0.05, 0.6), values[k] * 0.62),
            fontsize=10, color=C["red"], fontweight="bold",
            arrowprops=dict(arrowstyle="->", color=C["red"], lw=1.4),
        )

    for i in range(1, years + 1):                       # 逐年数据标注（年限少时）
        if years <= 12:
            ax.annotate(f"{values[i]:,.0f}", (t[i], values[i]),
                        textcoords="offset points", xytext=(0, 7),
                        ha="center", fontsize=8, color=C["dark"])

    ax.set_xlabel("年数")
    ax.set_ylabel("金额（元）")
    ax.set_title(f"{title}\n年利率 {rate:.1%} · {'年' if freq == 1 else f'每年计息 {freq} 次'}复利", loc="left")
    ax.yaxis.set_major_formatter(FuncFormatter(_money))
    ax.grid(axis="y", linestyle="--", alpha=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.legend(loc="upper left", fontsize=10)

    # ---------- 右：多利率对照 ----------
    ax2 = fig.add_subplot(gs[0, 1])
    rates = list(compare_rates) if compare_rates else [rate, rate + 0.02, rate + 0.04]
    rates = sorted({round(x, 6) for x in rates if x > -1})
    ends = [pv * (1.0 + x / freq) ** (freq * years) for x in rates]
    bars = ax2.barh([f"{x:.1%}" for x in rates], ends,
                    color=[C["primary"] if abs(x - rate) < 1e-9 else C["gray"] for x in rates],
                    height=0.52)
    for b, val in zip(bars, ends):
        ax2.text(val * 1.01, b.get_y() + b.get_height() / 2, f"{val:,.0f}",
                 va="center", fontsize=9, color=C["dark"])
    ax2.set_xlabel(f"{years} 年后的本息合计（元）")
    ax2.set_title("不同利率下的终值对照", loc="left")
    ax2.xaxis.set_major_formatter(FuncFormatter(_money))
    ax2.grid(axis="x", linestyle="--", alpha=0.8)
    ax2.set_axisbelow(True)
    ax2.set_xlim(0, max(ends) * 1.18)
    for s in ("top", "right"):
        ax2.spines[s].set_visible(False)

    fig.savefig(out_path)
    plt.close(fig)
    return out_path


# =============================================================================
# 图 2：年金现金流时间轴
# =============================================================================
def plot_annuity_timeline(
    pmt: float,
    rate: float,
    periods: int,
    out_path: str,
    due: bool = False,
    present_value: Optional[float] = None,
    future_value: Optional[float] = None,
    title: str = "年金现金流时间轴",
) -> str:
    """
    绘制年金现金流时间轴（stem 图）+ 折现贡献分解。

    上图：每期现金流以竖向箭头表示（向上=收入，向下=支出），期初/期末位置区分
          普通年金与预付年金；虚线标注现值/终值时点。
    下图：各期现金流折现到 0 时点的现值贡献，展示"远期现金流的现值衰减"，
          这是年金现值系数非线性的直观来源。
    """
    _ensure_dir(out_path)
    n = int(periods)
    fig, (ax, ax2) = plt.subplots(2, 1, figsize=(12.5, 6.6),
                                  gridspec_kw={"height_ratios": [1.35, 1.0], "hspace": 0.42})

    offset = 0.0 if due else 0.5          # 预付年金画在期初刻度，普通年金画在期中
    xs = np.arange(1, n + 1) - offset
    ys = np.full(n, float(pmt))

    # ---------- 上图：现金流轴 ----------
    ax.axhline(0, color=C["dark"], linewidth=1.6)
    markerline, stemlines, baseline = ax.stem(xs, ys, basefmt=" ")
    plt.setp(stemlines, color=C["primary"], linewidth=2.4, alpha=0.9)
    plt.setp(markerline, color=C["primary"], markersize=7)

    for x, y in zip(xs, ys):
        ax.annotate(f"{y:,.0f}", (x, y), textcoords="offset points",
                    xytext=(0, 9), ha="center", fontsize=8.5, color=C["dark"])
        # 折现缩放示意：用小灰点表示折现后的"实际价值感"
        ax.scatter([x], [y * (1 + rate) ** (-x)], color=C["gray"], s=16, zorder=4)

    ax.set_ylim(min(0, pmt * 1.35), pmt * 1.45 if pmt > 0 else pmt * 0.35)
    ax.set_xticks(np.arange(0, n + 1, 1 if n <= 15 else max(1, n // 12)))
    ax.set_xlabel("期数（时间轴）")
    ax.set_ylabel("现金流（元）")
    ax.set_title(
        f"{title} · {'预付年金（期初收付）' if due else '普通年金（期末收付）'} · 每期 {pmt:,.0f} 元 · 期利率 {rate:.2%}",
        loc="left")
    ax.grid(axis="y", linestyle="--", alpha=0.75)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)

    # 现值/终值时点标注（放在时间轴下方，避免与标题/柱形重叠）
    ax.set_ylim(-pmt * 1.05, pmt * 1.45)
    if present_value is not None:
        ax.scatter([0], [0], color=C["accent"], s=110, marker="D", zorder=6, clip_on=False)
        ax.annotate(f"现值 PVA\n{present_value:,.0f} 元", xy=(0, 0),
                    xytext=(0.15, -pmt * 0.72), ha="left", va="top", fontsize=9.5,
                    color=C["accent"], fontweight="bold")
    if future_value is not None:
        ax.scatter([n], [0], color=C["green"], s=110, marker="D", zorder=6, clip_on=False)
        ax.annotate(f"终值 FVA\n{future_value:,.0f} 元", xy=(n, 0),
                    xytext=(n - 0.15, -pmt * 0.72), ha="right", va="top", fontsize=9.5,
                    color=C["green"], fontweight="bold")

    # ---------- 下图：折现贡献 ----------
    contrib = np.array([pmt / (1 + rate) ** x for x in xs])
    ax2.bar(xs, contrib, width=0.55, color=C["green"], alpha=0.85, label="折现到 0 时点的现值贡献")
    if present_value is not None:
        ax2.axhline(present_value, color=C["accent"], linestyle="--", linewidth=1.6,
                    label=f"年金现值合计 {present_value:,.0f} 元")
    ax2.set_xlabel("期数")
    ax2.set_ylabel("现值贡献（元）")
    ax2.set_title("各期现金流折现贡献：越远期的现金流，现值衰减越快", loc="left")
    ax2.set_xticks(ax.get_xticks())
    ax2.yaxis.set_major_formatter(FuncFormatter(_money))
    ax2.grid(axis="y", linestyle="--", alpha=0.75)
    ax2.set_axisbelow(True)
    ax2.legend(loc="upper right", fontsize=9.5)
    for s in ("top", "right"):
        ax2.spines[s].set_visible(False)

    fig.savefig(out_path)
    plt.close(fig)
    return out_path


# =============================================================================
# 图 3：WACC 资本结构与成本分解
# =============================================================================
def plot_wacc_breakdown(
    w: Dict[str, float],
    out_path: str,
    industry_avg: float = 0.10,
    title: str = "WACC 资本结构与成本分解",
) -> str:
    """
    左：资本结构环形图（股权/债权/优先股占比）
    中：各资本来源成本对比（股权成本 vs 税前债权 vs 税后债权）
    右：加权贡献瀑布式条形，直观展示"谁在推高 WACC"
    """
    _ensure_dir(out_path)
    fig, axs = plt.subplots(1, 3, figsize=(15, 5.0),
                            gridspec_kw={"width_ratios": [1.0, 1.05, 1.1], "wspace": 0.30})

    we, wd, wp = w["股权比重"], w["债权比重"], w["优先股比重"]

    # ---------- 左：环形图 ----------
    ax = axs[0]
    labels, sizes, colors = [], [], []
    for lbl, val, col in (("股权 E", we, C["primary"]), ("债权 D", wd, C["accent"]), ("优先股 P", wp, C["purple"])):
        if val > 1e-9:
            labels.append(f"{lbl}\n{val:.1%}")
            sizes.append(val)
            colors.append(col)
    ax.pie(sizes, labels=labels, colors=colors, autopct="",
           startangle=90, wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
           textprops=dict(fontsize=10.5))
    ax.text(0, 0.06, "资本结构", ha="center", va="center", fontsize=11.5, color=C["gray"])
    ax.text(0, -0.12, f"总资本\n{w['资本总额']:,.0f}", ha="center", va="center",
            fontsize=10, color=C["dark"], fontweight="bold")
    ax.set_title("资本结构（市场价值口径）", loc="center")

    # ---------- 中：各来源成本对比 ----------
    ax = axs[1]
    kd_pre = w.get("_税前债权成本", w["税后债权成本"])
    names = ["股权成本\nKs", "债权成本\n(税前)", "债权成本\n(税后)"]
    vals = [w["股权成本"], kd_pre, w["税后债权成本"]]
    cols = [C["primary"], C["gray"], C["accent"]]
    bars = ax.bar(names, vals, color=cols, width=0.55)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v * 1.03, f"{v:.2%}",
                ha="center", fontsize=10.5, color=C["dark"], fontweight="bold")

    # 用箭头标示"税盾效应"带来的成本下降
    ax.annotate("", xy=(2, w["税后债权成本"]), xytext=(1, kd_pre),
                arrowprops=dict(arrowstyle="->", color=C["green"], lw=2, ls="--"))
    ax.text(1.5, (kd_pre + w["税后债权成本"]) / 2 * 0.96, "税盾效应", ha="center",
            fontsize=9.5, color=C["green"], fontweight="bold")

    if wp > 1e-9 and w["优先股成本"] > 0:
        ax.axhline(w["优先股成本"], color=C["purple"], linestyle=":", linewidth=1.5,
                   label=f"优先股成本 {w['优先股成本']:.2%}")
        ax.legend(fontsize=9)
    ax.set_ylabel("资本成本率")
    ax.yaxis.set_major_formatter(FuncFormatter(_pct))
    ax.set_ylim(0, max(vals) * 1.30)
    ax.set_title("各资本来源的成本水平", loc="center")
    ax.grid(axis="y", linestyle="--", alpha=0.75)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)

    # ---------- 右：加权贡献 ----------
    ax = axs[2]
    comps = [("股权", we * w["股权成本"], C["primary"]),
             ("债权", wd * w["税后债权成本"], C["accent"])]
    if wp > 1e-9:
        comps.append(("优先股", wp * w["优先股成本"], C["purple"]))
    comps.sort(key=lambda x: -x[1])

    bottom = 0.0
    for name, val, col in comps:
        ax.bar([0], [val], bottom=bottom, color=col, width=0.42,
               label=f"{name} 贡献 {val:.2%}")
        ax.text(0.30, bottom + val / 2, f"{name}\n{val:.2%}", va="center",
                fontsize=10, color=C["dark"])
        bottom += val
    ax.axhline(w["WACC"], color=C["red"], linestyle="--", linewidth=1.8)
    ax.annotate(f"WACC = {w['WACC']:.2%}", xy=(0.42, w["WACC"]),
                xytext=(0.62, w["WACC"] * 1.08), fontsize=11, fontweight="bold", color=C["red"])
    ax.axhline(industry_avg, color=C["gray"], linestyle=":", linewidth=1.5)
    ax.annotate(f"行业平均 {industry_avg:.1%}", xy=(0.42, industry_avg),
                xytext=(0.62, industry_avg * 1.02), fontsize=9.5, color=C["gray"])
    ax.set_xlim(-0.35, 1.1)
    ax.set_xticks([])
    ax.set_ylabel("对 WACC 的贡献（百分点）")
    ax.yaxis.set_major_formatter(FuncFormatter(_pct))
    ax.set_title("加权贡献累计", loc="center")
    ax.legend(loc="upper left", fontsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.75)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)

    fig.suptitle(title, fontsize=14.5, fontweight="bold", x=0.012, ha="left", y=1.02)
    fig.savefig(out_path)
    plt.close(fig)
    return out_path


def _shield_ratio(w: Dict[str, float]) -> float:
    """由税后成本反推税前成本时使用：(1 - 税率) 的近似——此处直接用 WACC 输入关系。"""
    k_after = w.get("税后债权成本", 0.0)
    k_pre = w.get("_税前债权成本", k_after)     # 由调用方注入
    return 0.0 if k_pre <= 0 else max(0.0, 1 - k_after / k_pre)


# =============================================================================
# 图 4：WACC 敏感性分析热力图
# =============================================================================
def plot_wacc_sensitivity(
    out_path: str,
    we: float,
    wd: float,
    base_ks: float,
    base_kd: float,
    ks_range: Tuple[float, float] = (-0.04, 0.04),
    kd_range: Tuple[float, float] = (-0.03, 0.03),
    steps: int = 9,
    industry_avg: float = 0.10,
    title: str = "WACC 敏感性分析：股权成本 vs 债权成本",
) -> str:
    """
    热力图展示 WACC 对 Ks 与 Kd 的联合敏感度，并在图上标注基准点与盈亏平衡线（WACC=行业平均）。
    用途：融资谈判中"债务利率上浮多少个百分点会击穿既定资本成本目标"的定量回答。
    """
    _ensure_dir(out_path)
    ks = np.linspace(base_ks + ks_range[0], base_ks + ks_range[1], steps)
    kd = np.linspace(base_kd + kd_range[0], base_kd + kd_range[1], steps)
    Z = np.array([[we * i + wd * j for i in ks] for j in kd])

    fig, ax = plt.subplots(figsize=(10.2, 6.6))
    # RdYlGn：低值=红（融资贵）、高值=绿（融资便宜）；颜色映射与"贵/便宜"语义一致
    mesh = ax.pcolormesh(ks, kd, Z, cmap="RdYlGn", shading="auto")
    cbar = fig.colorbar(mesh, ax=ax, pad=0.02)
    cbar.set_label("WACC", rotation=270, labelpad=16)
    cbar.ax.yaxis.set_major_formatter(FuncFormatter(_pct))

    for i in range(steps):                              # 数值标注
        for j in range(steps):
            ax.text(ks[i], kd[j], f"{Z[j, i]:.2%}", ha="center", va="center",
                    fontsize=8.0, color=C["dark"])

    base_w = we * base_ks + wd * base_kd
    ax.scatter([base_ks], [base_kd], marker="*", s=360, color="white",
               edgecolor=C["dark"], linewidth=1.6, zorder=5)
    ax.annotate(f"当前基准情形  WACC {base_w:.2%}",
                xy=(base_ks, base_kd), xytext=(ks[0] + 0.004, kd[-1] - 0.006),
                fontsize=10, fontweight="bold", color=C["dark"], zorder=7,
                bbox=dict(boxstyle="round,pad=0.30", fc="white", ec=C["dark"], alpha=0.92),
                arrowprops=dict(arrowstyle="->", color=C["dark"], lw=1.4))

    # 行业平均 WACC 的等值线 Kd = (WACC_industry - wE*Ks) / wD
    iso = (industry_avg - we * ks) / wd
    inside = (iso >= kd[0]) & (iso <= kd[-1])
    if inside.any():
        ax.plot(ks, iso, color="white", linestyle="--", linewidth=1.8, zorder=4)
        xi = int(np.argmax(inside))
        ax.annotate(f"WACC = 行业平均 {industry_avg:.0%}",
                    xy=(ks[xi], iso[xi]), xytext=(ks[xi] - 0.004, iso[xi] + 0.0092),
                    fontsize=9, color=C["dark"], fontweight="bold", zorder=6,
                    bbox=dict(boxstyle="round,pad=0.22", fc="white", ec=C["gray"], alpha=0.9))

    # 显式设定坐标范围，避免 pcolormesh 收缩绘图区导致大面积留白；
    # 顶部额外留出 0.6 格边距，防止最上一行的数值标注与坐标轴刻度文字重叠
    ax.set_xlim(ks[0] - (ks[1] - ks[0]) * 0.6, ks[-1] + (ks[1] - ks[0]) * 0.6)
    ax.set_ylim(kd[0] - (kd[1] - kd[0]) * 0.6, kd[-1] + (kd[1] - kd[0]) * 2.4)
    # 坐标轴刻度用百分比，并隐藏与格内数字重复的首末刻度标签（避免重叠）
    ax.set_xticks(ks)
    ax.set_yticks(kd)
    ax.xaxis.set_major_formatter(FuncFormatter(_pct))
    ax.yaxis.set_major_formatter(FuncFormatter(_pct))
    ax.tick_params(labelsize=9)
    ax.set_xticklabels([f"{x:.0%}" if i % 2 == 0 else "" for i, x in enumerate(ks)])
    ax.set_yticklabels([f"{y:.0%}" if i % 2 == 0 else "" for i, y in enumerate(kd)])
    ax.set_xlabel("股权资本成本 Ks")
    ax.set_ylabel("税后债权资本成本 Kd")
    ax.set_title(f"{title}\n（绿色 = WACC 低、融资成本便宜；红色 = WACC 高、融资成本昂贵）",
                 loc="left", fontsize=13)
    fig.savefig(out_path)
    plt.close(fig)
    return out_path


# =============================================================================
# 图 5：多方案 WACC 对比（资本结构优化）
# =============================================================================
def plot_wacc_scenarios(
    scenarios: List[Dict[str, float]],
    out_path: str,
    title: str = "资本结构优化：不同负债水平下的 WACC",
) -> str:
    """
    对比不同负债比例下的 WACC，标示最优点（WACC 最低的资本结构）。
    对应第 4 章"资本结构决策"知识点：适度负债降低 WACC，但过度负债
    会因财务风险上升推高 Kd 与 Ks，形成 U 型曲线。
    """
    _ensure_dir(out_path)
    names = [s["名称"] for s in scenarios]
    vals = [s["WACC"] for s in scenarios]
    best = int(np.argmin(vals))

    fig, ax = plt.subplots(figsize=(11.5, 5.4))
    bars = ax.bar(names, vals,
                  color=[C["green"] if i == best else C["primary"] for i in range(len(vals))],
                  width=0.55)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v * 1.02, f"{v:.3%}",
                ha="center", fontsize=10.5, fontweight="bold", color=C["dark"])

    ax.plot(names, vals, color=C["accent"], marker="o", linewidth=2.2,
            markersize=7, zorder=5, label="WACC 变化路径")
    ax.scatter([names[best]], [vals[best]], marker="*", s=340, color=C["red"],
               zorder=6, label=f"最优资本结构：{names[best]}（WACC 最低）")

    ax.set_ylabel("WACC")
    ax.yaxis.set_major_formatter(FuncFormatter(lambda x, p: f"{x:.2%}"))
    ax.set_ylim(min(vals) * 0.92, max(vals) * 1.10)
    ax.set_title(title, loc="left")
    ax.grid(axis="y", linestyle="--", alpha=0.75)
    ax.set_axisbelow(True)
    ax.legend(loc="upper center", ncol=2, fontsize=9.5)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.savefig(out_path)
    plt.close(fig)
    return out_path


# =============================================================================
# 图 6：债券资本成本敏感性（发行价 → YTM）
# =============================================================================
def plot_debt_cost_curve(
    out_path: str,
    face_value: float,
    coupon_rate: float,
    years: int,
    prices: Optional[Sequence[float]] = None,
    tax_rate: float = 0.25,
    title: str = "债权资本成本曲线：发行价与税后成本的关系",
) -> str:
    """
    绘制"发行价 → 税前/税后债务资本成本"曲线，标注平价发行点与税盾效应。
    直观解释：折价发行（价格低于面值）意味着投资者要求更高收益率，
    企业债务成本上升；同时两条曲线之间的垂直距离即为税盾带来的节省。
    """
    from core import cost_of_debt

    _ensure_dir(out_path)
    if prices is None:
        prices = np.linspace(face_value * 0.80, face_value * 1.10, 31)
    k_pre, k_post = [], []
    for p in prices:
        d = cost_of_debt(face_value, coupon_rate, years, float(p), tax_rate)
        k_pre.append(d["税前年化资本成本"])
        k_post.append(d["税后年化资本成本"])

    fig, ax = plt.subplots(figsize=(11.5, 5.4))
    ax.plot(prices, k_pre, color=C["primary"], linewidth=2.4, label="税前债务资本成本")
    ax.plot(prices, k_post, color=C["accent"], linewidth=2.4, label="税后债务资本成本")
    ax.fill_between(prices, k_post, k_pre, color=C["green"], alpha=0.18,
                    label=f"税盾效应节省（所得税率 {tax_rate:.0%}）")

    ax.axvline(face_value, color=C["gray"], linestyle="--", linewidth=1.4)
    y_min, y_max = min(k_post), max(k_pre)
    ax.annotate(f"平价发行 {face_value:,.0f} 元\n（税前成本 = 票面利率 {coupon_rate:.2%}）",
                xy=(face_value, coupon_rate), xytext=(face_value - (max(prices) - min(prices)) * 0.40,
                                                     y_min + (y_max - y_min) * 0.06),
                fontsize=9.5, color=C["gray"], ha="center",
                bbox=dict(boxstyle="round,pad=0.28", fc="white", ec=C["gray"], alpha=0.9),
                arrowprops=dict(arrowstyle="->", color=C["gray"], lw=1.2))
    ax.axhline(coupon_rate, color=C["gray"], linestyle=":", linewidth=1.3)
    ax.annotate(f"票面利率 {coupon_rate:.2%}", xy=(max(prices), coupon_rate),
                xytext=(max(prices), coupon_rate * 1.03), ha="right",
                fontsize=9.5, color=C["gray"])

    ax.set_xlabel("发行价 / 市价（元）")
    ax.set_ylabel("债务资本成本（年化）")
    ax.xaxis.set_major_formatter(FuncFormatter(_money))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda x, p: f"{x:.2%}"))
    ax.set_title(title, loc="left")
    ax.grid(linestyle="--", alpha=0.75)
    ax.set_axisbelow(True)
    ax.legend(loc="upper right", fontsize=10)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.savefig(out_path)
    plt.close(fig)
    return out_path
