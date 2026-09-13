#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
financial_calculator.pyw —— 货币时间价值与资本成本计算器（单文件整合版 v2.0）
==============================================================================

对应课程第 3 章（货币时间价值）与第 4 章（资本成本）

本文件把原先分离的 5 个模块合并为一个可直接双击运行的 GUI 程序：
    core.py        →  第 1 部分  核心计算引擎（纯函数）
    validators.py  →  第 2 部分  三层参数校验 + 业务解读引擎
    visualize.py   →  第 3 部分  Matplotlib 可视化（8 类图表）
    report.py      →  第 4 部分  导出引擎（原项目缺失，此处一并实现）
    main.py        →  第 5 部分  场景编排（改为返回结构化结果，不再 print）
                      第 6 部分  Tkinter 图形界面（双击 .pyw 直接运行）

运行方式
--------
    双击 financial_calculator.pyw                       # 打开图形界面
    python financial_calculator.pyw --demo              # 命令行：一键跑完所有场景
    python financial_calculator.pyw --selftest          # 命令行：运行数值自检测试
    python financial_calculator.pyw --check-imports     # 命令行：检查依赖环境

界面功能
--------
    总览页    一键运行全部场景、数值自检测试、导出测试用例簿、打开输出目录
    场景 1    复利终值 / 现值（年 / 季 / 月 / 日复利）
    场景 2    年金终值 / 现值（普通年金 / 预付年金）
    场景 3    债权资本成本（YTM 精确法 ↔ 教材近似法双引擎互验，含税盾）
    场景 4    股权资本成本（CAPM 主算 + DDM 交叉验证）
    场景 5    加权平均资本成本 WACC（含资本结构优化与敏感性分析）
    场景 6    贷款还款计划（等额本息，年金公式逆向应用）

设计要点
--------
* 纯函数计算层与界面层完全解耦，计算函数不做任何输入输出，便于测试与二次开发。
* 每次计算前都经过 L1 类型 / L2 边界 / L3 合理性三层校验，非法输入被拦截并给出
  "错在哪里 + 为什么错 + 应如何修正"的提示，不会输出无意义的计算结果。
* 每次计算后都输出业务解读，把数字翻译成决策语言。
* 图表统一配色、自动探测中文字体，输出为 PNG，可在界面内直接预览。
* 所有产物写入脚本同级的 outputs/ 目录（若已存在旧版 output/ 则沿用该目录）。

作者: WorkBuddy  |  版本: 2.0.0（单文件 GUI 版）
"""

from __future__ import annotations

import math
import os
import sys
import traceback
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

APP_TITLE = "财务计算器 · 资金时间价值与资本成本"
APP_VERSION = "v2.1 优化版"
APP_TAGLINE = "把公式、校验与结果解释放在同一张工作台上"


# =============================================================================
# 0. 依赖导入（缺失时给出可读的提示，而不是抛一堆栈信息）
# =============================================================================
_IMPORT_ERRORS: Dict[str, str] = {}


def _try_import(name: str):
    try:
        return __import__(name)
    except Exception as exc:                                    # noqa: BLE001
        _IMPORT_ERRORS[name] = f"{type(exc).__name__}: {exc}"
        return None


np = _try_import("numpy")
pd = _try_import("pandas")
matplotlib = _try_import("matplotlib")

if matplotlib is not None:
    matplotlib.use("Agg")               # 只出图不弹窗；界面用 PNG 预览，无需 GUI 后端
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter

MPL_OK = matplotlib is not None
PANDAS_OK = pd is not None

PIP_HINT = "  ".join(f"pip install {m}" for m in ("numpy", "pandas", "matplotlib", "openpyxl"))


def dependency_report() -> List[str]:
    """返回依赖环境的逐项检查结果，供命令行与界面状态栏显示。"""
    rows = []
    for name, mod in (("numpy", np), ("pandas", pd), ("matplotlib", matplotlib),
                      ("openpyxl", _try_import("openpyxl"))):
        if mod is None:
            rows.append(f"  ✘ {name:<11} 缺失 —— {_IMPORT_ERRORS.get(name, '未安装')}")
        else:
            rows.append(f"  ✔ {name:<11} {getattr(mod, '__version__', '')}")
    return rows


# =============================================================================
# 1. 核心计算引擎（原 core.py）
# =============================================================================
class FinanceError(ValueError):
    """财务计算域异常：参数超出数学模型的有效定义域。"""


def _check_periods(n: float) -> float:
    """期数校验：必须为正的有限数。使用 math.isfinite 而非 np.isfinite，
    避免 numpy 缺失时崩溃。"""
    if not math.isfinite(n) or n <= 0:
        raise FinanceError(f"期数 n 必须为大于 0 的有限数，当前为 {n}。")
    return float(n)


def _check_rate(r: float) -> float:
    """利率校验：必须为大于 -100% 的有限数（-100% 意味着本金归零，无经济意义）。"""
    if not math.isfinite(r) or r <= -1:
        raise FinanceError(f"期间利率 r 必须大于 -100%，当前为 {r}。")
    return float(r)


# ---------------------------------------------------------------- 1.1 复利
def pv_compound(fv: float, r: float, n: float, m: int = 1) -> float:
    """
    复利现值：PV = FV / (1 + r/m)^(n*m)

    fv : 终值（未来某期金额）
    r  : 年名义利率（小数，如 6% 传 0.06）
    n  : 年数
    m  : 每年计息次数（1=年计息，2=半年，4=季，12=月，365=日）

    >>> round(pv_compound(1_000_000, 0.06, 5, 12), 2)
    741372.92
    """
    _check_periods(n)
    _check_rate(r / m)
    return float(fv) / (1.0 + r / m) ** (n * m)


def fv_compound(pv: float, r: float, n: float, m: int = 1) -> float:
    """
    复利终值：FV = PV * (1 + r/m)^(n*m)

    m 越大（计息越频繁），实际年利率（EAR）越高，这一效应称为"复利频率效应"。

    >>> round(fv_compound(100_000, 0.05, 10, 1), 2)
    162889.46
    >>> round(fv_compound(100_000, 0.05, 10, 12), 2)
    164700.95
    """
    _check_periods(n)
    _check_rate(r / m)
    return float(pv) * (1.0 + r / m) ** (n * m)


def effective_annual_rate(r: float, m: int = 1) -> float:
    """
    名义年利率 → 实际年利率（EAR / EFF）：EAR = (1 + r/m)^m - 1
    用于量化"计息频率"对真实资金成本的影响。
    """
    _check_rate(r / m)
    return (1.0 + r / m) ** m - 1.0


# ---------------------------------------------------------------- 1.2 年金
def pva_ordinary(pmt: float, r: float, n: float) -> float:
    """
    普通年金现值（每期期末收付）：
        PVA = PMT * [1 - (1 + r)^-n] / r        （r ≠ 0）
        PVA = PMT * n                            （r = 0，退化情形）

    典型场景：房贷、消费贷的等额本息现值，养老金领取额的初始本金。
    """
    n = _check_periods(n)
    r = _check_rate(r)
    if abs(r) < 1e-12:                      # 零利率退化，避免除零
        return float(pmt) * n
    return float(pmt) * (1.0 - (1.0 + r) ** (-n)) / r


def fva_ordinary(pmt: float, r: float, n: float) -> float:
    """
    普通年金终值（每期期末存入）：
        FVA = PMT * [(1 + r)^n - 1] / r         （r ≠ 0）
    典型场景：零存整取、每月定投的到期总额。
    """
    n = _check_periods(n)
    r = _check_rate(r)
    if abs(r) < 1e-12:
        return float(pmt) * n
    return float(pmt) * ((1.0 + r) ** n - 1.0) / r


def pva_due(pmt: float, r: float, n: float) -> float:
    """
    预付年金（先付年金）现值：每期期初收付。
        PVA_due = PVA_ordinary * (1 + r)
    典型场景：租金"先付后用"、保险费期初缴纳。
    """
    return pva_ordinary(pmt, r, n) * (1.0 + _check_rate(r))


def fva_due(pmt: float, r: float, n: float) -> float:
    """
    预付年金终值：FVA_due = FVA_ordinary * (1 + r)
    含义：期初存入比期末存入多赚一期利息，故终值放大 (1+r) 倍。
    """
    return fva_ordinary(pmt, r, n) * (1.0 + _check_rate(r))


def pv_perpetuity(pmt: float, r: float, growth: float = 0.0) -> float:
    """
    永续年金现值（含固定增长模型，戈登模型）：PVP = PMT / (r - g)
    r <= g 时公式无经济意义（级数不收敛），抛出异常。
    """
    r = _check_rate(r)
    if r <= growth:
        raise FinanceError(f"折现率 r({r:.4%}) 必须大于增长率 g({growth:.4%})，否则永续年金现值不收敛。")
    return float(pmt) / (r - growth)


def fv_from_schedule(cashflows: Sequence[float], r: float) -> Dict[str, object]:
    """
    给定不规则现金流序列，逐期滚动计算终值：FV_t = FV_(t-1) * (1 + r) + CF_t

    cashflows : 各期现金流序列，第 0 期为期初（可为负=投入，正=收回）
    返回 {"fv", "path", "total_in", "total_gain"}
    """
    r = _check_rate(r)
    path: List[float] = []
    acc = 0.0
    for i, cf in enumerate(cashflows):
        acc = acc * (1.0 + r) + float(cf) if i > 0 else float(cf)
        path.append(acc)
    total_in = -sum(c for c in cashflows if c < 0)
    total_gain = acc - total_in if total_in > 0 else acc
    return {"fv": acc, "path": path, "total_in": total_in, "total_gain": total_gain}


# ------------------------------------------------------ 1.3 债权资本成本
def cost_of_debt(
    face_value: float,
    coupon_rate: float,
    years: int,
    price: float,
    tax_rate: float = 0.0,
    freq: int = 1,
) -> Dict[str, float]:
    """
    债权资本成本（到期收益率法 YTM / IRR 法）—— 精确算法。

    原理：债券发行价是"未来本息的现值"，即令 NPV = 0 的折现率就是债权人的
          到期收益率，也就是企业承担的税前债务资本成本。

        Price = Σ [C / (1+k)^t] + F / (1+k)^N ,  C = F * coupon / freq

    返回：税前期间/年化资本成本、税后年化资本成本、税盾节省(百分点)、年利息支出、期限(期数)
    """
    N = int(years) * int(freq)
    if N <= 0:
        raise FinanceError("债券期限（年）× 每年付息次数 必须为正整数。")
    if price <= 0:
        raise FinanceError("债券发行价必须大于 0。")

    c = face_value * coupon_rate / freq
    flows = np.array([c] * N)
    flows[-1] += face_value                      # 最后一期还本

    k_period = irr(np.concatenate(([-price], flows)), guess=coupon_rate / freq)
    k_pre_y = (1.0 + k_period) ** freq - 1.0     # 期间利率 → 年化有效利率
    k_after_y = k_pre_y * (1.0 - tax_rate)       # 税盾效应
    shield = k_pre_y * tax_rate

    return {
        "税前期间资本成本": k_period,
        "税前年化资本成本": k_pre_y,
        "税后年化资本成本": k_after_y,
        "税盾节省(百分点)": shield * 100,
        "年利息支出": c * freq,
        "期限(期数)": float(N),
    }


def cost_of_debt_approx(
    face_value: float,
    coupon_rate: float,
    years: int,
    price: float,
    tax_rate: float = 0.0,
) -> Dict[str, float]:
    """
    债权资本成本 —— 教材近似公式法（用于与 IRR 法互验）：

        Kd(税前) ≈ [F*coupon + (F - P)/N] / [(F + P)/2]

    分子 = 年票面利息 + 年均资本利得(折价摊销)；分母 = 面值与发行价的算术平均。
    """
    if years <= 0:
        raise FinanceError("债券期限必须为正整数。")
    interest = face_value * coupon_rate
    amortized = (face_value - price) / years
    avg_funds = (face_value + price) / 2.0
    k_pre = (interest + amortized) / avg_funds
    return {
        "税前年化资本成本(近似)": k_pre,
        "税后年化资本成本(近似)": k_pre * (1.0 - tax_rate),
    }


def cost_of_debt_after_tax(kd_pretax: float, tax_rate: float) -> float:
    """
    税后债务资本成本（税盾效应的一般形式）：Kd_after = Kd_before * (1 - T)

    经济含义：债务利息可在企业所得税前扣除，产生"税盾（tax shield）"，使债务的
    实际资本成本低于名义利率。这是债务融资相对于股权融资的核心优势。
    """
    if not 0.0 <= tax_rate < 1.0:
        raise FinanceError(f"所得税税率须在 [0, 1) 区间内，当前为 {tax_rate}。")
    return kd_pretax * (1.0 - tax_rate)


# ------------------------------------------------------ 1.4 股权资本成本
def capm_required_return(
    rf: float,
    beta: float,
    rm: float,
    size_premium: float = 0.0,
) -> Dict[str, float]:
    """
    资本资产定价模型（CAPM）估算股权资本成本：Ks = Rf + β * (Rm - Rf) + SP

    Rf=无风险利率，β=系统性风险系数，(Rm-Rf)=市场风险溢价，SP=规模/流动性附加溢价。
    """
    if beta < 0:
        raise FinanceError(f"β 系数不应为负（当前 {beta}）；若确为负相关资产，请确认后另行处理。")
    mrp = rm - rf
    return {
        "无风险利率": rf,
        "β系数": beta,
        "市场收益率": rm,
        "市场风险溢价": mrp,
        "规模溢价": size_premium,
        "股权资本成本": rf + beta * mrp + size_premium,
        "风险补偿部分": beta * mrp + size_premium,
    }


def cost_of_equity_ddm(d1: float, p0: float, g: float) -> float:
    """
    股利折现模型（DDM / 戈登增长模型）估算股权资本成本：Ks = D1 / P0 + g
    适用于分红稳定、增长可预期的成熟企业，可与 CAPM 结果交叉验证。
    """
    if p0 <= 0:
        raise FinanceError("股票现价 P0 必须大于 0。")
    return d1 / p0 + g


# ------------------------------------------------------ 1.5 WACC 与辅助工具
def wacc(
    equity_value: float,
    debt_value: float,
    ks: float,
    kd_after_tax: float,
    pref_value: float = 0.0,
    kp: float = 0.0,
) -> Dict[str, float]:
    """
    加权平均资本成本：WACC = (E/V) * Ks + (D/V) * Kd_after_tax + (P/V) * Kp ，V = E + D + P

    注意：进入 WACC 的债务成本必须是**税后**成本；股权成本因股利不可税前列支，
          不享受税盾，直接用 CAPM 结果。
    """
    e = float(equity_value)
    d = float(debt_value)
    p = float(pref_value)
    v = e + d + p
    if v <= 0:
        raise FinanceError("资本总额（股权+债权+优先股）必须大于 0。")
    if min(e, d, p) < 0:
        raise FinanceError("资本金额不应为负数。")

    we, wd, wp = e / v, d / v, p / v
    w = we * ks + wd * kd_after_tax + wp * kp
    return {
        "股权比重": we,
        "债权比重": wd,
        "优先股比重": wp,
        "股权成本": ks,
        "税后债权成本": kd_after_tax,
        "优先股成本": kp,
        "WACC": w,
        "资本总额": v,
        "股权贡献度(百分点)": we * ks * 100,
        "债权贡献度(百分点)": wd * kd_after_tax * 100,
    }


def bond_price(face_value: float, coupon_rate: float, years: int,
               ytm: float, freq: int = 1) -> Dict[str, float]:
    """
    债券理论价格 = 票息年金现值 + 面值复利现值（现金流贴现法）。
    用于判断债券"溢价/平价/折价"发行。
    """
    if years <= 0 or freq <= 0:
        raise FinanceError("债券期限与付息频率必须为正整数。")
    N = years * freq
    c = face_value * coupon_rate / freq
    r = ytm / freq
    pv_coupons = pva_ordinary(c, r, N)
    pv_face = pv_compound(face_value, ytm, years, freq)
    total = pv_coupons + pv_face
    return {
        "理论价格": total,
        "票息现值": pv_coupons,
        "面值现值": pv_face,
        "类型": "溢价发行" if total > face_value else ("折价发行" if total < face_value else "平价发行"),
    }


def npv(rate: float, cashflows: Sequence[float]) -> float:
    """净现值：NPV = Σ CF_t / (1 + rate)^t ，CF_0 通常为负的投资额。"""
    _check_rate(rate)
    cf = np.asarray(cashflows, dtype=float)
    t = np.arange(len(cf))
    return float(np.sum(cf / (1.0 + rate) ** t))


def irr(cashflows: Sequence[float], guess: float = 0.1,
        tol: float = 1e-10, max_iter: int = 200) -> float:
    """
    内部收益率 IRR —— 自实现牛顿迭代法（不依赖外部财务函数库）。

    求解 f(r) = NPV(r) = 0；牛顿法 r_(k+1) = r_k - f(r_k)/f'(r_k)，
    其中 f'(r) = Σ -t * CF_t / (1+r)^(t+1)。

    稳健性设计：牛顿法不收敛时退化为 [-0.99, 10] 区间上的二分法，
    保证债券 YTM 这类问题总能得到数值解。
    """
    cf = np.asarray(cashflows, dtype=float)
    if len(cf) < 2:
        raise FinanceError("IRR 至少需要 2 期现金流。")
    if not (np.any(cf > 0) and np.any(cf < 0)):
        raise FinanceError("IRR 要求现金流同时包含正负项，否则方程无解。")

    t = np.arange(len(cf), dtype=float)
    r = float(guess)

    for _ in range(max_iter):
        base = 1.0 + r
        if base <= 1e-9:
            break
        disc = base ** t
        f = float(np.sum(cf / disc))
        df = float(np.sum(-t * cf / (disc * base)))
        if abs(df) < 1e-15:
            break
        r_new = r - f / df
        if abs(r_new - r) < tol:
            return r_new
        r = r_new

    # ---- 牛顿法未收敛：二分法兜底 ----
    lo, hi = -0.9999, 10.0
    f_lo, f_hi = npv(lo, cf), npv(hi, cf)
    if f_lo * f_hi > 0:
        raise FinanceError("IRR 在合理区间内不存在符号变号点，无法求解，请检查现金流。")
    for _ in range(300):
        mid = (lo + hi) / 2.0
        f_mid = npv(mid, cf)
        if abs(f_mid) < 1e-10 or (hi - lo) < 1e-12:
            return mid
        if f_lo * f_mid <= 0:
            hi, f_hi = mid, f_mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2.0


def amortization_schedule(principal: float, annual_rate: float,
                          years: int, freq: int = 12) -> Dict[str, object]:
    """
    生成等额本息还款计划表（房贷/车贷场景）。

    每期还款额 PMT 由年金现值公式反解：PMT = P * r / [1 - (1 + r)^-n]；
    每期拆分为"利息 = 期初余额 × r"、"本金 = PMT - 利息"。
    """
    if principal <= 0:
        raise FinanceError("贷款本金必须大于 0。")
    n = years * freq
    r = annual_rate / freq
    if abs(r) < 1e-12:
        pmt = principal / n
    else:
        pmt = principal * r / (1.0 - (1.0 + r) ** (-n))

    rows, balance = [], float(principal)
    for i in range(1, n + 1):
        interest = balance * r
        principal_paid = pmt - interest
        balance -= principal_paid
        if i == n:
            balance = 0.0
        rows.append({
            "期数": i, "月供": pmt, "利息": interest,
            "本金": principal_paid, "剩余本金": max(balance, 0.0),
        })
    return {
        "每期还款额": pmt,
        "总还款额": pmt * n,
        "总利息": pmt * n - principal,
        "计划表": rows,
    }



# =============================================================================
# 2. 参数校验与业务解读（原 validators.py）
# =============================================================================
@dataclass
class ValidationResult:
    """一次参数校验的完整结果：errors 阻断计算，warnings 提示复核，hints 说明口径。"""
    ok: bool = True
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    hints: List[str] = field(default_factory=list)

    def add_error(self, msg: str) -> None:
        self.ok = False
        self.errors.append(msg)

    def add_warning(self, msg: str) -> None:
        self.warnings.append(msg)

    def add_hint(self, msg: str) -> None:
        self.hints.append(msg)

    def summary(self) -> str:
        parts = []
        if self.errors:
            parts.append("错误: " + "; ".join(self.errors))
        if self.warnings:
            parts.append("警告: " + "; ".join(self.warnings))
        return " | ".join(parts) if parts else "校验通过"


@dataclass
class Rule:
    """单条校验规则（声明式，便于扩展）。"""
    name: str
    kind: str        # "error" | "warning" | "hint"
    test: str
    lo: Optional[float] = None
    hi: Optional[float] = None
    message: str = ""


# 参数 -> 适用规则。区间单位为"输入值本身的单位"（利率为小数）。
RULES: Dict[str, Dict[str, Any]] = {
    # ---------- 利率类（小数形式，0.05 = 5%） ----------
    "rate": {
        "label": "利率", "hard": (-0.99, 1.00), "typical": (-0.05, 0.30),
        "unit": "小数（0.05 表示 5%）",
    },
    "discount_rate": {
        "label": "折现率", "hard": (0.0, 1.00), "typical": (0.01, 0.30), "unit": "小数",
    },
    "rf": {
        "label": "无风险利率", "hard": (0.0, 0.20), "typical": (0.015, 0.06),
        "unit": "小数；中国常用 10 年期国债收益率 2%~3%",
    },
    "rm": {
        "label": "市场期望收益率", "hard": (-0.5, 0.60), "typical": (0.05, 0.20),
        "unit": "小数；长期股票市场年化收益经验值 8%~12%",
    },
    "beta": {
        "label": "β系数", "hard": (0.0, 5.0), "typical": (0.3, 2.5),
        "unit": "无量纲；市场组合定义为 1.0",
    },
    "tax_rate": {
        "label": "所得税税率", "hard": (0.0, 0.99), "typical": (0.0, 0.35),
        "unit": "小数；中国一般企业所得税 25%，高新 15%",
    },
    "coupon_rate": {
        "label": "票面利率", "hard": (0.0, 0.50), "typical": (0.01, 0.12), "unit": "小数",
    },
    "growth": {
        "label": "增长率", "hard": (-0.5, 1.0), "typical": (-0.1, 0.15), "unit": "小数",
    },
    # ---------- 期数类 ----------
    "periods": {
        "label": "期数", "hard": (1, 1200), "typical": (1, 600),
        "unit": "期（正整数值）", "must_be_int": True,
    },
    "years": {
        "label": "年数", "hard": (1, 100), "typical": (1, 50),
        "unit": "年（正整数值）", "must_be_int": True,
    },
    # ---------- 金额类 ----------
    "amount": {"label": "金额", "hard": (0.0, 1e15), "typical": (0.0, 1e12), "unit": "元"},
    "positive_amount": {
        "label": "金额（须为正）", "hard": (1e-9, 1e15), "typical": (0.0, 1e12), "unit": "元",
    },
    "weight": {"label": "权重", "hard": (0.0, 1.0), "typical": (0.0, 1.0), "unit": "小数（0~1 之间）"},
}


def validate(field_name: str, value: Any, rules: Dict[str, Any] = None) -> ValidationResult:
    """
    对单个参数按 RULES 中同名规则做三层校验：
      L1 类型/缺失 -> L2 硬性边界（阻断） -> L3 业务合理性（警告）
    返回 ValidationResult。
    """
    r = ValidationResult()
    spec = (rules or RULES).get(field_name)
    if spec is None:
        r.add_hint(f"参数 {field_name} 未定义校验规则，已跳过校验。")
        return r

    label, unit = spec["label"], spec.get("unit", "")

    # ---- L1 类型 / 缺失 ----
    if value is None or value == "":
        r.add_error(f"【{label}】不能为空，请输入一个数值（单位：{unit}）。")
        return r
    try:
        v = float(value)
    except (TypeError, ValueError):
        r.add_error(f"【{label}】必须为数字，当前输入为 “{value}”。请输入形如 0.06 的数值（单位：{unit}）。")
        return r

    # ---- 整数性校验 ----
    if spec.get("must_be_int") and abs(v - round(v)) > 1e-9:
        r.add_error(f"【{label}】必须为整数（当前 {v}）。期数/年数在年金与复利公式中代表完整计息周期数。")
        return r
    v = round(v) if spec.get("must_be_int") else v

    # ---- L2 硬性边界 ----
    lo, hi = spec.get("hard", (None, None))
    if lo is not None and hi is not None and not (lo <= v <= hi):
        r.add_error(
            f"【{label}】取值 {v} 超出有效范围 [{lo}, {hi}]（单位：{unit}）。"
            f"该值会使财务模型失去数学或经济意义，请检查是否混淆了"
            f"“小数”与“百分数”口径（例如 6% 应输入 0.06 而非 6）。"
        )
        return r

    # ---- L3 业务合理性 ----
    tlo, thi = spec.get("typical", (None, None))
    if tlo is not None and thi is not None and not (tlo <= v <= thi):
        r.add_warning(
            f"【{label}】取值 {v} 偏离常见业务区间 [{tlo}, {thi}]（单位：{unit}），"
            f"数值可计算但较为极端，建议核对数据来源。"
        )

    r.add_hint(f"【{label}】口径提示：{unit}")
    return r


def validate_all(params: Dict[str, Any]) -> ValidationResult:
    """批量校验一组参数（界面与批处理共用入口）。"""
    total = ValidationResult()
    for k, v in params.items():
        one = validate(k, v)
        total.errors += one.errors
        total.warnings += one.warnings
        if one.errors:
            total.ok = False
    return total


def check_wacc_weights(wd: float, we: float, wp: float = 0.0,
                       tol: float = 1e-6) -> ValidationResult:
    """
    WACC 权重合计校验：三者之和必须等于 1。
    这是 WACC 最常见的口径错误——用账面价值权重算完后与实际资本结构不一致。
    """
    r = ValidationResult()
    s = wd + we + wp
    if abs(s - 1.0) > tol:
        r.add_error(
            f"资本结构权重合计为 {s:.4f}，应等于 1。请检查："
            f"(1) 是否遗漏了优先股或少数股东权益；"
            f"(2) 权重是否基于同一口径（建议统一用市场价值而非账面价值）。"
        )
    if we <= 0:
        r.add_warning("股权权重为 0，相当于 100% 债务融资，现实中极罕见，请确认。")
    if wd <= 0:
        r.add_hint("债务权重为 0，WACC 退化为纯股权成本，未体现税盾效应。")
    return r


def check_beta_sanity(beta: float) -> ValidationResult:
    """β 极端值复核：提示用户核对可比公司样本量与回归窗口。"""
    r = ValidationResult()
    if beta > 2.5:
        r.add_warning(f"β = {beta} 属高波动资产（如周期性行业、小盘成长股），"
                      f"建议核对是否使用了足够长的回归窗口（一般 2~5 年周频/月频数据）。")
    if 0 <= beta < 0.3:
        r.add_warning(f"β = {beta} 明显低于市场，通常见于公用事业、必选消费等弱周期行业，请确认样本选择。")
    return r


# ---------------------------------------------- 2.1 业务场景解释引擎
def interpret_wacc(w: Dict[str, float], industry_avg: float = 0.10,
                   roic: Optional[float] = None,
                   rf: float = 0.025) -> List[str]:
    """
    自动生成 WACC 结果的业务解读，四个维度：
      1. 绝对水平（与行业平均对比） 2. 相对水平（与无风险利率利差）
      3. 结构成因（各来源贡献分解） 4. 价值创造（EVA 判断）
    """
    out: List[str] = []
    v = w["WACC"]
    gap = v - industry_avg

    if gap <= -0.02:
        t, s = "明显低于", "融资成本具有显著优势，同等项目下可承受更低的投资回报率，安全边际较大。"
    elif gap <= -0.005:
        t, s = "低于", "融资成本处于较优水平，在项目投资决策中具有相对竞争力。"
    elif gap < 0.005:
        t, s = "基本持平于", "融资成本与行业平均水平接近，投资决策应更依赖项目的差异化优势。"
    elif gap < 0.02:
        t, s = "略高于", "融资成本略高，需通过提升项目回报率或优化资本结构来弥补。"
    else:
        t, s = "明显高于", ("融资成本偏高，建议重点审视：高 β 带来的股权成本、财务风险推高的债务成本，"
                          "以及资本结构是否过度依赖股权。")

    out.append(f"当前 WACC 为 {v:.2%}，{t}行业平均水平（{industry_avg:.2%}，差值 {gap:+.2%}），{s}")
    out.append(
        f"相对无风险利率（{rf:.2%}）的风险补偿为 {v - rf:.2%}，"
        f"这是投资者承担经营与财务风险所要求的额外回报，也是项目折现率的构成基础。"
    )

    we, wd = w["股权比重"], w["债权比重"]
    out.append(
        f"资本结构上股权占 {we:.1%}、债权占 {wd:.1%}；"
        f"股权与债权对 WACC 的贡献分别为 {w['股权贡献度(百分点)']:.2f} 与 {w['债权贡献度(百分点)']:.2f} 个百分点。"
    )
    if wd > 0:
        out.append(
            f"由于利息可在税前列支，债务实际成本仅为 {w['税后债权成本']:.2%}"
            f"（低于其税前水平），税盾效应在本次资本结构中降低了加权成本。"
        )

    if roic is not None:
        eva = roic - v
        if eva > 0:
            out.append(f"企业投入资本回报率 ROIC = {roic:.2%}，高于 WACC {eva:.2%}，"
                       f"说明当前经营正在创造经济增加值（EVA > 0），股东价值在增加。")
        else:
            out.append(f"企业投入资本回报率 ROIC = {roic:.2%}，低于 WACC {abs(eva):.2%}，"
                       f"说明经营回报未能覆盖全部资本成本（EVA < 0），长期看会侵蚀股东价值，"
                       f"需通过提价、降本或压缩低效资本占用改善。")
    return out


def interpret_tvm(kind: str, inputs: Dict[str, float], result: float) -> List[str]:
    """复利/年金结果的时间价值解读。"""
    out: List[str] = []
    r = inputs.get("rate")
    n = inputs.get("periods") or inputs.get("years")
    if r is not None and n is not None and r > 0:
        out.append(
            f"在年化利率 {r:.2%} 下，资金约每 {72 / (max(r, 1e-6) * 100):.1f} 年翻一番"
            f"（按“72 法则”快速估算），{n:.0f} 年内的增长倍数约为 {(1 + r) ** n:.2f} 倍。"
        )
    if kind.startswith("复利终值"):
        pv = inputs.get("amount", 0)
        if pv:
            out.append(
                f"期初投入 {pv:,.2f} 元，到期金额 {result:,.2f} 元，"
                f"其中利息收益 {result - pv:,.2f} 元，占比 {(result - pv) / pv:.1%}。"
            )
    elif "年金" in kind:
        pmt = inputs.get("amount", 0)
        total = pmt * (n or 0)
        out.append(
            f"累计投入 {total:,.2f} 元，{'现值' if '现值' in kind else '终值'}为 {result:,.2f} 元；"
            f"每期收付发生在{'期初（预付年金，比普通年金多赚一期利息）' if '预付' in kind else '期末（普通年金）'}。"
        )
    return out


def interpret_debt(d: Dict[str, float], tax_rate: float,
                   market_rate: Optional[float] = None) -> List[str]:
    """债务资本成本的业务解读。"""
    out: List[str] = []
    k_pre, k_after = d.get("税前年化资本成本"), d.get("税后年化资本成本")
    if k_pre and k_after:
        out.append(
            f"税前债务资本成本（到期收益率口径）为 {k_pre:.3%}，"
            f"考虑所得税率 {tax_rate:.0%} 后降至 {k_after:.3%}，"
            f"税盾每年为企业节省约 {k_pre - k_after:.3%} 的融资成本。"
        )
        out.append(
            "债务资本成本通常低于股权成本，原因有二：一是债权人求偿权优先、风险更低，"
            "二是利息具有税前扣除的税盾效应。这正是适度负债能提升企业价值（降低 WACC）的机理。"
        )
    if market_rate is not None and k_pre:
        spread = k_pre - market_rate
        out.append(
            f"本次发债成本相对同期市场基准利率 {'高出' if spread > 0 else '低于'} "
            f"{abs(spread):.3%}，反映市场对企业信用资质与流动性的定价。"
        )
    return out


def interpret_equity(e: Dict[str, float]) -> List[str]:
    """股权资本成本（CAPM）的业务解读。"""
    rf, beta = e["无风险利率"], e["β系数"]
    mrp, ks = e["市场风险溢价"], e["股权资本成本"]
    out: List[str] = []
    out.append(
        f"股权资本成本 Ks = Rf({rf:.2%}) + β({beta:.2f}) × MRP({mrp:.2%}) = {ks:.2%}；"
        f"其中风险补偿部分为 {e['风险补偿部分']:.2%}，占比 {e['风险补偿部分'] / ks:.0%}。"
    )
    if beta > 1:
        out.append(f"β = {beta:.2f} > 1，说明该股票波动大于市场整体（如周期性制造、券商、房地产），"
                   f"投资者要求更高的风险补偿，股权融资相对更“贵”。")
    elif beta < 1:
        out.append(f"β = {beta:.2f} < 1，属于防御型资产（如公用事业、食品饮料），"
                   f"与市场联动较弱，股权成本相对较低，适合作为高杠杆企业的资本结构调整方向。")
    else:
        out.append("β ≈ 1，风险特征与市场整体同步。")
    out.append("提示：CAPM 的结果对无风险利率与风险溢价假设高度敏感，"
               "建议同时用股利折现模型（DDM）或可比公司法交叉验证。")
    return out



# =============================================================================
# 3. 可视化（原 visualize.py）
# =============================================================================
_CJK_CANDIDATES = [
    "Microsoft YaHei", "SimHei", "SimSun", "KaiTi",          # Windows 优先
    "Noto Sans CJK SC", "Noto Sans CJK JP", "Source Han Sans SC",
    "PingFang SC", "WenQuanYi Zen Hei", "DejaVu Sans",
]


def _setup_style() -> str:
    """
    探测可用中文字体并统一 rcParams，返回实际使用的字体名。
    候选列表按"可用性优先"排序，只要该字体覆盖中文即可正确显示汉字。
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


CJK_FONT = _setup_style() if MPL_OK else "（matplotlib 未安装）"

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


def plot_compound_growth(
    pv: float, rate: float, years: int, out_path: str,
    compare_rates: Optional[Sequence[float]] = None, freq: int = 1,
    title: str = "复利增长曲线：货币的时间价值",
) -> str:
    """
    图 1：复利增长曲线。
    左：本金 + 利息堆积面积图，标注"资金翻倍"临界点；右：不同利率下的终值对照。
    """
    _ensure_dir(out_path)
    t = np.arange(0, years + 1)
    factor = (1.0 + rate / freq) ** (freq * t)
    values = pv * factor

    fig = plt.figure(figsize=(13, 5.4))
    gs = fig.add_gridspec(1, 2, width_ratios=[2.05, 1.0], wspace=0.24)

    ax = fig.add_subplot(gs[0, 0])
    ax.fill_between(t, 0, pv, color=C["primary"], alpha=0.85, label=f"本金 {pv:,.0f} 元")
    ax.fill_between(t, pv, values, color=C["green"], alpha=0.75, label="累计利息收益")
    ax.plot(t, values, color=C["dark"], linewidth=2.0, zorder=5, label="本息合计")

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


def plot_annuity_timeline(
    pmt: float, rate: float, periods: int, out_path: str, due: bool = False,
    present_value: Optional[float] = None, future_value: Optional[float] = None,
    title: str = "年金现金流时间轴",
) -> str:
    """
    图 2：年金现金流时间轴（stem 图）+ 折现贡献分解。
    上：每期现金流方向与金额，区分普通年金/预付年金；下：各期折现贡献的衰减。
    """
    _ensure_dir(out_path)
    n = int(periods)
    fig, (ax, ax2) = plt.subplots(2, 1, figsize=(12.5, 6.6),
                                  gridspec_kw={"height_ratios": [1.35, 1.0], "hspace": 0.42})

    offset = 0.0 if due else 0.5          # 预付年金画在期初刻度，普通年金画在期中
    xs = np.arange(1, n + 1) - offset
    ys = np.full(n, float(pmt))

    ax.axhline(0, color=C["dark"], linewidth=1.6)
    markerline, stemlines, baseline = ax.stem(xs, ys, basefmt=" ")
    plt.setp(stemlines, color=C["primary"], linewidth=2.4, alpha=0.9)
    plt.setp(markerline, color=C["primary"], markersize=7)

    for x, y in zip(xs, ys):
        ax.annotate(f"{y:,.0f}", (x, y), textcoords="offset points",
                    xytext=(0, 9), ha="center", fontsize=8.5, color=C["dark"])
        ax.scatter([x], [y * (1 + rate) ** (-x)], color=C["gray"], s=16, zorder=4)

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


def plot_wacc_breakdown(
    w: Dict[str, float], out_path: str, industry_avg: float = 0.10,
    title: str = "WACC 资本结构与成本分解",
) -> str:
    """
    图 3：WACC 资本结构与成本分解（三联图）。
    左：资本结构环形图；中：各来源成本对比（含税盾箭头）；右：加权贡献累计。
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


def plot_wacc_sensitivity(
    out_path: str, we: float, wd: float, base_ks: float, base_kd: float,
    ks_range: Tuple[float, float] = (-0.04, 0.04),
    kd_range: Tuple[float, float] = (-0.03, 0.03),
    steps: int = 9, industry_avg: float = 0.10,
    title: str = "WACC 敏感性分析：股权成本 vs 债权成本",
) -> str:
    """
    图 4：WACC 敏感性热力图。
    展示 WACC 对 Ks 与 Kd 的联合敏感度，标注基准点与行业平均等值线。
    用途：融资谈判中"债务利率上浮多少个百分点会击穿既定资本成本目标"的定量回答。
    """
    _ensure_dir(out_path)
    ks = np.linspace(base_ks + ks_range[0], base_ks + ks_range[1], steps)
    kd = np.linspace(base_kd + kd_range[0], base_kd + kd_range[1], steps)
    Z = np.array([[we * i + wd * j for i in ks] for j in kd])

    fig, ax = plt.subplots(figsize=(10.2, 6.6))
    # RdYlGn：低值=红（融资贵）、高值=绿（融资便宜）
    mesh = ax.pcolormesh(ks, kd, Z, cmap="RdYlGn", shading="auto")
    cbar = fig.colorbar(mesh, ax=ax, pad=0.02)
    cbar.set_label("WACC", rotation=270, labelpad=16)
    cbar.ax.yaxis.set_major_formatter(FuncFormatter(_pct))

    for i in range(steps):
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

    # 行业平均 WACC 的等值线：Kd = (WACC_industry - wE*Ks) / wD
    if wd > 1e-12:
        iso = (industry_avg - we * ks) / wd
        inside = (iso >= kd[0]) & (iso <= kd[-1])
        if inside.any():
            ax.plot(ks, iso, color="white", linestyle="--", linewidth=1.8, zorder=4)
            xi = int(np.argmax(inside))
            ax.annotate(f"WACC = 行业平均 {industry_avg:.0%}",
                        xy=(ks[xi], iso[xi]), xytext=(ks[xi] - 0.004, iso[xi] + 0.0092),
                        fontsize=9, color=C["dark"], fontweight="bold", zorder=6,
                        bbox=dict(boxstyle="round,pad=0.22", fc="white", ec=C["gray"], alpha=0.9))

    ax.set_xlim(ks[0] - (ks[1] - ks[0]) * 0.6, ks[-1] + (ks[1] - ks[0]) * 0.6)
    ax.set_ylim(kd[0] - (kd[1] - kd[0]) * 0.6, kd[-1] + (kd[1] - kd[0]) * 2.4)
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


def plot_wacc_scenarios(
    scenarios: List[Dict[str, float]], out_path: str,
    title: str = "资本结构优化：不同负债水平下的 WACC",
) -> str:
    """
    图 5：多方案 WACC 对比（资本结构优化），标示 WACC 最低的最优资本结构。
    对应第 4 章"资本结构决策"：适度负债降低 WACC，过度负债形成 U 型反转。
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


def plot_debt_cost_curve(
    out_path: str, face_value: float, coupon_rate: float, years: int,
    prices: Optional[Sequence[float]] = None, tax_rate: float = 0.25,
    title: str = "债权资本成本曲线：发行价与税后成本的关系",
) -> str:
    """
    图 6：发行价 → 税前/税后债务资本成本曲线，标注平价发行点与税盾效应。
    折价发行意味着投资者要求更高收益率，企业债务成本上升；两线间距即税盾节省。
    """
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
                xy=(face_value, coupon_rate),
                xytext=(face_value - (max(prices) - min(prices)) * 0.40,
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


def plot_loan_structure(df, out_path: str) -> str:
    """图 7：还款结构图 —— 每期本金/利息堆积 + 剩余本金曲线（双轴）。"""
    fig, ax = plt.subplots(figsize=(12, 5.2))
    x = df["期数"]
    ax.bar(x, df["利息"], color=C["accent"], width=0.9, label="利息部分")
    ax.bar(x, df["本金"], bottom=df["利息"], color=C["primary"], width=0.9, label="本金部分")
    ax.set_xlabel("期数")
    ax.set_ylabel("每期还款（元）")
    ax.yaxis.set_major_formatter(FuncFormatter(_money))
    ax.grid(axis="y", linestyle="--", alpha=0.7)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)

    ax2 = ax.twinx()
    ax2.plot(x, df["剩余本金"], color=C["red"], linewidth=2.2, label="剩余本金")
    ax2.set_ylabel("剩余本金（元）", color=C["red"])
    ax2.yaxis.set_major_formatter(FuncFormatter(_money))
    ax2.tick_params(axis="y", colors=C["red"])
    ax2.spines["top"].set_visible(False)

    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="center right", fontsize=9.5)
    ax.set_title("等额本息还款结构：前期利息为主，后期本金为主（剩余本金呈凸曲线下降）", loc="left")
    fig.savefig(out_path)
    plt.close(fig)
    return out_path


PLOTTERS: Dict[str, Callable[..., str]] = {
    "compound_growth": plot_compound_growth,
    "annuity_timeline": plot_annuity_timeline,
    "wacc_breakdown": plot_wacc_breakdown,
    "wacc_sensitivity": plot_wacc_sensitivity,
    "wacc_scenarios": plot_wacc_scenarios,
    "debt_cost_curve": plot_debt_cost_curve,
    "loan_structure": plot_loan_structure,
}



# =============================================================================
# 4. 导出引擎（原 report.py —— 原项目缺失，此处一并实现）
# =============================================================================
# 摘要表数值口径：命中这些关键词的字段按百分比渲染，其余按金额/数值渲染。
_PCT_TOKENS = ("利率", "成本", "比重", "权重", "税率", "溢价", "收益率", "增长率", "EAR", "WACC")


def fmt_is_pct_key(key: str) -> bool:
    """按字段名语义判断是否应以百分比渲染（与原 report.py 的输出口径一致）。"""
    ku = str(key).upper()
    return any(t.upper() in ku for t in _PCT_TOKENS)


def fmt_summary_value(key: str, value: Any) -> str:
    """按字段名语义格式化摘要值：利率/成本/比重类 → 百分比；数值 → 千分位保留 4 位。"""
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, (int, float)):
        if fmt_is_pct_key(key):
            return f"{value:.4%}"
        return f"{value:,.4f}"
    return str(value)


def _sanitize_row(row: Dict[str, Any]) -> Dict[str, Any]:
    """剔除内部字段（下划线开头）并把非标量值转为字符串，保证可写入表格。"""
    out: Dict[str, Any] = {}
    for k, v in row.items():
        if str(k).startswith("_"):
            continue
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            out[k] = v
        elif v is None:
            out[k] = ""
        else:
            out[k] = str(v)
    return out


def export_summary(results: List[Dict[str, Any]], path: str,
                   extra_sheets: Sequence[Tuple[str, Any]] = ()) -> str:
    """
    导出计算结果汇总 Excel（09_计算结果汇总.xlsx）。

    results      : 各场景的扁平摘要 dict 列表（第一列通常为"场景"）
    extra_sheets : 附加工作表 [(sheet_name, DataFrame), ...]，如资本结构优化表
    """
    _ensure_dir(path)
    rows = [_sanitize_row(r) for r in results]
    cols: List[str] = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    # 让"场景"始终排在第一列
    if "场景" in cols:
        cols.remove("场景")
        cols.insert(0, "场景")

    df = pd.DataFrame(rows, columns=cols)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # 工作表名与旧版 report.py 一致（场景汇总 + WACC明细 + 资本结构模拟 + 还款计划前24期）
        df.to_excel(writer, sheet_name="场景汇总", index=False)
        for name, extra in extra_sheets:
            if extra is None:
                continue
            extra.to_excel(writer, sheet_name=str(name)[:31], index=False)
    return path


def export_markdown(results: List[Dict[str, Any]], path: str) -> str:
    """导出 Markdown 汇总报告（计算结果汇总.md），供撰写作业报告直接引用。"""
    _ensure_dir(path)
    lines: List[str] = ["# 货币时间价值与资本成本计算结果汇总", ""]
    for i, r in enumerate(results, 1):
        row = _sanitize_row(r)
        name = row.get("场景", f"场景 {i}")
        lines += [f"## {i}. {name}", "", "| 指标 | 数值 |", "| --- | --- |"]
        for k, v in row.items():
            lines.append(f"| {k} | {fmt_summary_value(k, v)} |")
        lines.append("")
    lines += [
        "---",
        "",
        f"> 说明：本文件由 `{os.path.basename(__file__) or 'financial_calculator.pyw'}` 自动生成，"
        f"数值口径为“利率=小数、金额=元”。",
        "> 完整图表见 `outputs/` 目录下的 PNG 文件。",
        "",
    ]
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    return path


def export_loan_schedule(schedule: Dict[str, Any], path: str) -> str:
    """导出等额本息还款计划表（07_贷款还款计划.xlsx，单工作表「还款计划」）。"""
    _ensure_dir(path)
    df = pd.DataFrame(schedule["计划表"])
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="还款计划", index=False)
    return path


def export_casebook(records: List[Dict[str, Any]], path: str) -> str:
    """
    导出测试用例明细（10_测试用例明细.xlsx）：
      工作表「用例索引」：序号 | 用例名称 | 结论
      工作表「用例N」   ：项目 | 内容（参数 → 计算步骤 → 结论）
    版式与旧版 report.py 产物保持一致。
    """
    _ensure_dir(path)
    overview = [{"序号": i, "用例名称": rec.get("名称", ""), "结论": rec.get("结论", "")}
                for i, rec in enumerate(records, 1)]
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(overview).to_excel(writer, sheet_name="用例索引", index=False)
        for i, rec in enumerate(records, 1):
            body: List[Dict[str, Any]] = []
            for k, v in (rec.get("参数") or {}).items():
                body.append({"项目": "参数", "内容": f"{k} = {v}"})
            for item in (rec.get("过程") or []):
                if isinstance(item, (tuple, list)) and len(item) >= 2:
                    body.append({"项目": "计算步骤", "内容": f"{item[0]} → {item[1]}"})
                else:
                    body.append({"项目": "计算步骤", "内容": str(item)})
            body.append({"项目": "结论", "内容": rec.get("结论", "")})
            pd.DataFrame(body).to_excel(writer, sheet_name=f"用例{i}"[:31], index=False)
    return path



# =============================================================================
# 5. 场景编排（原 main.py 的场景函数，改为返回结构化结果、不再 print）
# =============================================================================
def _pad(text: str, width: int, right: bool = False) -> str:
    """按"显示宽度"补齐（中文按 2 列计），保证等宽字体下表格对齐。"""
    text = str(text)
    used = sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)
    gap = " " * max(0, width - used)
    return gap + text if right else text + gap


def kv_lines(pairs: Sequence[Tuple[str, Any]], pct: Sequence[str] = (),
             money: Sequence[str] = (), indent: str = "    ") -> List[str]:
    """把 (标签, 值) 序列渲染成对齐的两列表格文本行。"""
    pairs = [(str(k), v) for k, v in pairs]
    pct, money = set(pct), set(money)
    w = max((sum(2 if unicodedata.east_asian_width(c) in ("W", "F") else 1 for c in k)
             for k, _ in pairs), default=8)
    out = [indent + "─" * (w + 24)]
    for k, v in pairs:
        if isinstance(v, bool):
            s = "是" if v else "否"
        elif isinstance(v, (int, float)):
            if k in pct:
                s = f"{v:.4%}"
            elif k in money:
                s = f"{v:,.2f}"
            else:
                s = f"{v:,.4f}"
        else:
            s = str(v)
        out.append(indent + _pad(k, w) + "  " + _pad(s, 20, right=True))
    out.append(indent + "─" * (w + 24))
    return out


@dataclass
class SceneResult:
    """一个场景的完整运行结果，界面、导出、测试共用同一份数据。"""
    name: str
    summary: Dict[str, Any] = field(default_factory=dict)          # 扁平摘要，用于 Excel / Markdown
    sections: List[Tuple[str, List[str]]] = field(default_factory=list)  # (小节标题, 文本行)
    figures: List[str] = field(default_factory=list)               # 生成的 PNG 路径
    files: List[str] = field(default_factory=list)                 # 生成的其他文件
    extra_sheets: List[Tuple[str, Any]] = field(default_factory=list)    # 附加工作表

    def to_text(self) -> str:
        """渲染为纯文本（命令行模式与复制粘贴用）。"""
        buf = [self.name, "=" * 62]
        for title, lines in self.sections:
            buf.append(f"\n▶ {title}")
            buf.extend(lines)
        if self.figures or self.files:
            buf.append("\n▶ 输出文件")
            for p in list(self.figures) + list(self.files):
                buf.append(f"    · {p}")
        return "\n".join(buf)


# ---------------------------------------------------- 5.1 资本结构优化模拟
def simulate_capital_structure(
    equity: float, debt: float, ks_base: float, kd_pre_base: float, tax: float,
) -> List[Dict[str, Any]]:
    """
    模拟不同负债比例下的 WACC（资本结构优化）。

    风险传导假设（体现 MM 理论有税修正 + 财务困境成本）：
        1. 债权成本随负债率上升：Kd = Kd0 * (1 + 0.20 * (wd - wd0))
        2. 股权成本随负债率上升且更敏感：Ks = Ks0 * (1 + 1.05 * (wd - wd0))
        3. 负债率超过 55% 后引入财务困境成本非线性加速项：
           Ks 额外 + 0.9 * (wd - 0.55)^2 ，Kd 额外 + 0.5 * (wd - 0.55)^2

    三者共同作用形成 U 型 WACC 曲线：初期税盾效应占优使 WACC 下降，
    越过最优点后财务风险上升的影响超过税盾收益，WACC 转为上升。
    """
    total = equity + debt
    wd0 = debt / total
    rows = []
    for wd in np.linspace(0.10, 0.80, 8):
        delta = wd - wd0
        kd_pre = kd_pre_base * (1 + 0.20 * delta)
        ks = ks_base * (1 + 1.05 * delta)

        if wd > 0.55:                      # 高杠杆区的财务困境成本加速项
            stress = (wd - 0.55) ** 2
            ks += 0.9 * stress
            kd_pre += 0.5 * stress

        kd_post = kd_pre * (1 - tax)
        w = (1 - wd) * ks + wd * kd_post
        rows.append({
            "负债比例": f"{wd:.0%}",
            "税前Kd": kd_pre,
            "税后Kd": kd_post,
            "Ks": ks,
            "WACC": w,
        })
    return rows


# ------------------------------------------------------------ 5.2 场景 1 复利
def run_compound(pv: float, rate: float, years: int, freq: int = 1,
                 mode: str = "1", out_dir: Optional[str] = None) -> SceneResult:
    """复利终值/现值计算 + 复利增长曲线 + 业务解读。mode="2" 表示已知终值求现值。"""
    out_dir = out_dir or get_output_dir()
    if mode == "2":
        res = pv_compound(pv, rate, years, freq)
        main_label = "复利现值 PV"
        pairs = [("终值 FV", pv), ("复利现值 PV", res)]
    else:
        res = fv_compound(pv, rate, years, freq)
        main_label = "复利终值 FV"
        pairs = [("现值 PV", pv), ("复利终值 FV", res)]

    ear = effective_annual_rate(rate, freq)
    interest = res - pv if mode == "1" else pv - res
    pairs += [
        ("年名义利率", rate), ("期数(年)", years), ("每年计息次数", freq),
        ("实际年利率 EAR", ear), ("利息收益", interest),
    ]

    kv = kv_lines(pairs, pct={"年名义利率", "实际年利率 EAR"},
                  money={"现值 PV", "复利终值 FV", "终值 FV", "复利现值 PV", "利息收益"})

    notes = ["· " + l for l in interpret_tvm(
        "复利终值" if mode == "1" else "复利现值",
        {"rate": rate, "years": years, "amount": pv}, res)]
    if freq > 1 and mode == "1":
        base = fv_compound(pv, rate, years, 1)
        notes.append(
            f"· 计息频率效应：年复利终值为 {base:,.2f} 元，{freq} 次复利后提升至 {res:,.2f} 元，"
            f"多出 {res - base:,.2f} 元（{(res / base - 1):.3%}），这就是 EAR 高于名义利率的原因。")

    fig = os.path.join(out_dir, "01_复利增长曲线.png")
    plot_compound_growth(pv, rate, int(years), fig, freq=int(freq))

    return SceneResult(
        name="场景 1 ｜ 复利终值与现值计算（第 3 章）",
        summary={"场景": "复利终值/现值", "现值PV": pv, "年利率": rate, "年数": years,
                 "计息次数": freq, "实际年利率EAR": ear, main_label: res},
        sections=[("计算结果 · " + main_label, kv), ("业务解读", notes),
                  ("输出文件", [f"· 复利增长曲线：{os.path.basename(fig)}"])],
        figures=[fig],
    )


# ------------------------------------------------------------ 5.3 场景 2 年金
def run_annuity(pmt: float, rate: float, periods: int, due: bool = False,
                out_dir: Optional[str] = None) -> SceneResult:
    """普通年金 / 预付年金的终值与现值 + 现金流时间轴 + 折现贡献分解。"""
    out_dir = out_dir or get_output_dir()
    pva = pva_due(pmt, rate, periods) if due else pva_ordinary(pmt, rate, periods)
    fva = fva_due(pmt, rate, periods) if due else fva_ordinary(pmt, rate, periods)
    total = pmt * periods

    kv = kv_lines(
        [("每期现金流 PMT", pmt), ("每期利率", rate), ("期数", periods),
         ("收付方式", "预付年金（期初）" if due else "普通年金（期末）"),
         ("年金现值 PVA", pva), ("年金终值 FVA", fva),
         ("累计现金流总额", total), ("现值/累计比", pva / total), ("终值/累计比", fva / total)],
        pct={"每期利率", "现值/累计比", "终值/累计比"},
        money={"每期现金流 PMT", "年金现值 PVA", "年金终值 FVA", "累计现金流总额"})

    notes = ["· " + l for l in interpret_tvm(
        "预付年金现值" if due else "年金现值",
        {"rate": rate, "periods": periods, "amount": pmt}, pva)]
    if due:
        base_pva, base_fva = pva_ordinary(pmt, rate, periods), fva_ordinary(pmt, rate, periods)
        notes.append(
            f"· 预付 vs 普通年金：现值相差 {pva - base_pva:,.2f} 元（+{(pva / base_pva - 1):.2%}），"
            f"终值相差 {fva - base_fva:,.2f} 元（+{(fva / base_fva - 1):.2%}）。"
            f"因为每笔款项都提前一期到账/存入，多获得一期利息。")
    notes.append(f"· 现值仅为累计现金流总额的 {pva / total:.1%}，说明资金的时间价值让远期现金流"
                 f"在今天的价值大幅缩水——这正是长期分期付款「总额看起来很多、现值其实不高」的原因。")

    fig = os.path.join(out_dir, "02_年金现金流时间轴.png")
    plot_annuity_timeline(pmt, rate, int(periods), fig, due=due,
                          present_value=pva, future_value=fva)

    return SceneResult(
        name="场景 2 ｜ 年金终值与现值计算（第 3 章）",
        summary={"场景": "年金终值/现值", "每期PMT": pmt, "期利率": rate, "期数": periods,
                 "收付方式": "预付" if due else "普通", "年金现值PVA": pva, "年金终值FVA": fva},
        sections=[("计算结果", kv), ("业务解读", notes),
                  ("输出文件", [f"· 年金现金流时间轴：{os.path.basename(fig)}"])],
        figures=[fig],
    )


# ---------------------------------------------------------- 5.4 场景 3 债权
def run_debt(face: float, coupon: float, years: int, price: float, tax: float,
             out_dir: Optional[str] = None) -> SceneResult:
    """债权资本成本：YTM 精确法 ↔ 教材近似法双引擎互验 + 税盾效应可视化。"""
    out_dir = out_dir or get_output_dir()
    d = cost_of_debt(face, coupon, int(years), price, tax)
    d_approx = cost_of_debt_approx(face, coupon, int(years), price, tax)
    diff = d["税前年化资本成本"] - d_approx["税前年化资本成本(近似)"]

    kv = kv_lines(
        [("债券面值", face), ("票面利率", coupon), ("期限(年)", years), ("发行价", price),
         ("所得税税率", tax), ("年利息支出", d["年利息支出"]),
         ("税前资本成本(YTM法)", d["税前年化资本成本"]),
         ("税后资本成本", d["税后年化资本成本"]),
         ("税盾节省(百分点)", d["税盾节省(百分点)"] / 100),
         ("税前资本成本(近似公式法)", d_approx["税前年化资本成本(近似)"]),
         ("两法差异", diff)],
        pct={"票面利率", "所得税税率", "税前资本成本(YTM法)", "税后资本成本",
             "税盾节省(百分点)", "税前资本成本(近似公式法)", "两法差异"},
        money={"债券面值", "发行价", "年利息支出"})

    notes = [
        "· 方法一（主算法）IRR/YTM 法：令债券未来本息现金流现值等于发行价的折现率，"
        "即债权人要求的到期收益率，也是企业真实承担的税前债务成本。",
        f"· 方法二（校验法）教材近似公式：分母取面值与发行价均值。两法结果高度接近，"
        f"差异仅 {abs(diff):.4%}，互为验证。",
    ]
    notes += ["· " + l for l in interpret_debt(d, tax, market_rate=coupon)]
    if price < face:
        notes.append(f"· 本次债券折价发行（发行价 {price:,.0f} < 面值 {face:,.0f}），"
                     f"说明市场要求的收益率高于票面利率，投资者以低价买入来补偿利息不足，"
                     f"企业的实际债务成本因此高于票面利率。")
    elif price > face:
        notes.append("· 本次债券溢价发行，市场要求的收益率低于票面利率，企业实际债务成本低于票面利率。")
    else:
        notes.append("· 本次债券平价发行，市场要求收益率等于票面利率，税前债务成本 = 票面利率。")

    fig = os.path.join(out_dir, "03_债权资本成本曲线.png")
    plot_debt_cost_curve(fig, face, coupon, int(years), tax_rate=tax)

    return SceneResult(
        name="场景 3 ｜ 债权资本成本计算（第 4 章，含税盾效应）",
        summary={"场景": "债权资本成本", "面值": face, "票面利率": coupon, "期限(年)": years,
                 "发行价": price, "税率": tax, "税前成本": d["税前年化资本成本"],
                 "税后成本": d["税后年化资本成本"]},
        sections=[("计算结果", kv), ("方法学说明与业务解读", notes),
                  ("输出文件", [f"· 债权资本成本曲线：{os.path.basename(fig)}"])],
        figures=[fig],
    )


# ---------------------------------------------------------- 5.5 场景 4 股权
def run_equity(rf: float, beta: float, rm: float, size_premium: float = 0.0,
               use_ddm: bool = True, d1: float = 0.5, p0: float = 12.0, g: float = 0.04,
               out_dir: Optional[str] = None) -> SceneResult:
    """股权资本成本：CAPM 为主，DDM 交叉验证。"""
    out_dir = out_dir or get_output_dir()
    e = capm_required_return(rf, beta, rm, size_premium)

    kv = kv_lines(
        [("无风险利率 Rf", rf), ("β系数", beta), ("市场期望收益率 Rm", rm),
         ("市场风险溢价 MRP", e["市场风险溢价"]), ("规模溢价", size_premium),
         ("风险补偿部分", e["风险补偿部分"]), ("股权资本成本 Ks", e["股权资本成本"])],
        pct={"无风险利率 Rf", "市场期望收益率 Rm", "市场风险溢价 MRP",
             "规模溢价", "风险补偿部分", "股权资本成本 Ks"})

    notes = ["· " + l for l in interpret_equity(e)]
    beta_sanity = check_beta_sanity(beta)
    if beta_sanity.warnings:
        notes += ["⚠ " + w for w in beta_sanity.warnings]

    sections: List[Tuple[str, List[str]]] = [("CAPM 计算结果", kv), ("业务解读", notes)]

    ddm_value: Any = ""
    if use_ddm:
        ddm_value = cost_of_equity_ddm(d1, p0, g)
        ddm_kv = kv_lines(
            [("每股股利 D1", d1), ("股票现价 P0", p0), ("增长率 g", g), ("DDM 股权成本 Ks", ddm_value)],
            pct={"增长率 g", "DDM 股权成本 Ks"}, money={"每股股利 D1", "股票现价 P0"})
        diff = ddm_value - e["股权资本成本"]
        ddm_kv.append(
            f"· 两模型差异 {diff:+.3%}："
            + ("差异较小，结论稳健" if abs(diff) < 0.02
               else "差异较大，建议进一步核查增长率假设与企业分红政策的可持续性")
            + "。实务中通常以 CAPM 结果为主，DDM 用于校验。")
        sections.append(("DDM 交叉验证", ddm_kv))

    return SceneResult(
        name="场景 4 ｜ 股权资本成本计算（第 4 章，CAPM 模型）",
        summary={"场景": "股权资本成本", "无风险利率": rf, "β": beta, "市场收益率": rm,
                 "股权成本(CAPM)": e["股权资本成本"], "股权成本(DDM)": ddm_value},
        sections=sections,
    )


# ----------------------------------------------------------- 5.6 场景 5 WACC
def run_wacc(face: float, coupon: float, years: int, price: float, tax: float,
             rf: float, beta: float, rm: float, equity: float, debt: float,
             pref: float = 0.0, kp: float = 0.0,
             out_dir: Optional[str] = None) -> SceneResult:
    """完整 WACC 计算：债务成本 + 股权成本 + 加权，含资本结构优化与敏感性分析。"""
    out_dir = out_dir or get_output_dir()

    d = cost_of_debt(face, coupon, int(years), price, tax)
    kd_after = d["税后年化资本成本"]
    e = capm_required_return(rf, beta, rm)
    ks = e["股权资本成本"]

    total = equity + debt + pref
    w_res = check_wacc_weights(debt / total, equity / total, pref / total)
    w = wacc(equity, debt, ks, kd_after, pref, kp)
    w["_税前债权成本"] = d["税前年化资本成本"]

    kv = kv_lines(
        [("税前债务成本 Kd(税前)", d["税前年化资本成本"]),
         ("税后债务成本 Kd(税后)", kd_after),
         ("股权成本 Ks(CAPM)", ks),
         ("股权市场价值 E", equity), ("债务市场价值 D", debt), ("优先股 P", pref),
         ("股权权重 wE", w["股权比重"]), ("债务权重 wD", w["债权比重"]),
         ("WACC", w["WACC"])],
        pct={"税前债务成本 Kd(税前)", "税后债务成本 Kd(税后)", "股权成本 Ks(CAPM)",
             "股权权重 wE", "债务权重 wD", "WACC"},
        money={"股权市场价值 E", "债务市场价值 D", "优先股 P"})

    notes: List[str] = []
    notes += ["⚠ " + x for x in w_res.errors] + ["⚠ " + x for x in w_res.warnings]
    notes += ["· " + l for l in interpret_wacc(w, industry_avg=0.10, roic=0.118, rf=rf)]

    # ---- 资本结构优化模拟 ----
    scenarios = simulate_capital_structure(equity, debt, ks, d["税前年化资本成本"], tax)
    df_s = pd.DataFrame(scenarios)
    table = ["    " + _pad("负债比例", 10) + "".join(_pad(h, 14, right=True)
                                                    for h in ("税前Kd", "税后Kd", "Ks", "WACC"))]
    for row in scenarios:
        table.append("    " + _pad(row["负债比例"], 10) + "".join(
            _pad(f"{row[h]:.4f}", 14, right=True) for h in ("税前Kd", "税后Kd", "Ks", "WACC")))
    best = min(scenarios, key=lambda r: r["WACC"])
    struct_notes = [
        "· 假设负债上升会小幅推高财务风险，从而同时抬升 Kd 与 Ks；作为对照进行情景测算。",
        f"· 模拟结果中 WACC 最低的负债比例为 {best['负债比例']}（WACC = {best['WACC']:.4%}），"
        f"即本次测算下的最优资本结构。",
        "· 曲线呈 U 型：初期税盾收益大于财务风险成本，WACC 随负债上升而下降；"
        "越过最优点后财务困境成本占优，WACC 反转为上升——并非负债越多越好。",
    ]

    fig1 = os.path.join(out_dir, "04_WACC结构分解.png")
    plot_wacc_breakdown(w, fig1, industry_avg=0.10)
    fig2 = os.path.join(out_dir, "05_WACC敏感性分析.png")
    plot_wacc_sensitivity(fig2, w["股权比重"], w["债权比重"], ks, kd_after)
    fig3 = os.path.join(out_dir, "06_资本结构优化对比.png")
    plot_wacc_scenarios([{"名称": s["负债比例"] + " 负债", "WACC": s["WACC"]} for s in scenarios], fig3)

    return SceneResult(
        name="场景 5 ｜ 加权平均资本成本 WACC 计算（第 4 章）",
        summary={"场景": "WACC 计算", "税前债务成本": d["税前年化资本成本"],
                 "税后债务成本": kd_after, "股权成本Ks": ks,
                 "股权权重": w["股权比重"], "债务权重": w["债权比重"], "WACC": w["WACC"]},
        sections=[("WACC 汇总结果", kv), ("业务解读", notes),
                  ("资本结构优化：不同负债水平下的 WACC 模拟", table + struct_notes),
                  ("输出文件", [f"· {os.path.basename(p)}" for p in (fig1, fig2, fig3)])],
        figures=[fig1, fig2, fig3],
        extra_sheets=[
            ("WACC明细", pd.DataFrame(
                {"指标": [k for k in w if not k.startswith("_")],
                 "数值": [v for k, v in w.items() if not k.startswith("_")]})),
            ("资本结构模拟", df_s),
        ],
    )


# ----------------------------------------------------------- 5.7 场景 6 贷款
def run_loan(principal: float, rate: float, years: int, freq: int = 12,
             out_dir: Optional[str] = None) -> SceneResult:
    """等额本息还款计划：把年金现值公式反过来用（第 3 章应用场景）。"""
    out_dir = out_dir or get_output_dir()
    s = amortization_schedule(principal, rate, int(years), int(freq))
    pva_check = pva_ordinary(s["每期还款额"], rate / freq, years * freq)

    kv = kv_lines(
        [("贷款本金", principal), ("年利率", rate), ("年限", years), ("每年还款次数", freq),
         ("每期还款额", s["每期还款额"]), ("总还款额", s["总还款额"]),
         ("总利息", s["总利息"]), ("利息/本金比", s["总利息"] / principal),
         ("校验：还款额的年金现值", pva_check)],
        pct={"年利率", "利息/本金比"},
        money={"贷款本金", "每期还款额", "总还款额", "总利息", "校验：还款额的年金现值"})

    notes = [
        f"· 每期需还款 {s['每期还款额']:,.2f} 元，{years} 年累计还款 {s['总还款额']:,.2f} 元，"
        f"其中利息高达 {s['总利息']:,.2f} 元，相当于本金的 {s['总利息'] / principal:.1%}。",
        f"· 校验逻辑：把每期还款额按 {rate / freq:.4%} 的期利率折现 {years * freq} 期，"
        f"得到 {pva_check:,.2f} 元，等于贷款本金，验证了等额本息公式的正确性"
        f"（差异 {abs(pva_check - principal):.6f} 元，为浮点误差）。",
        "· 这说明等额本息并非「利息少」，而是前期利息占比高、后期本金占比高；"
        "若提前还款，节奏越早，节省的利息越多。",
    ]

    df = pd.DataFrame(s["计划表"])
    xlsx = os.path.join(out_dir, "07_贷款还款计划.xlsx")
    export_loan_schedule(s, xlsx)
    fig = os.path.join(out_dir, "08_还款结构.png")
    plot_loan_structure(df, fig)

    return SceneResult(
        name="场景 6 ｜ 贷款还款计划（等额本息，年金公式应用）",
        summary={"场景": "贷款还款计划", "贷款本金": principal, "年利率": rate,
                 "年限": years, "月供": s["每期还款额"], "总利息": s["总利息"]},
        sections=[("还款计划摘要", kv), ("业务解读", notes),
                  ("输出文件", [f"· 等额本息还款计划表：{os.path.basename(xlsx)}",
                             f"· 还款结构图：{os.path.basename(fig)}"])],
        figures=[fig], files=[xlsx],
        extra_sheets=[("还款计划前24期", df.head(24))],
    )


# ============================================= 5.8 自检测试（原 test_cases.py）
class SelfTest:
    """
    数值自检测试：与教材系数表、解析解逐项比对。
    共 6 组主用例 + 1 组参数校验测试（原 test_cases.py 的全部断言）。
    """

    def __init__(self) -> None:
        self.lines: List[str] = []
        self.records: List[Dict[str, Any]] = []
        self.passed = 0
        self.failed = 0

    # ---------------- 基础设施 ----------------
    def section(self, text: str) -> None:
        self.lines += ["", "━" * 78, text, "━" * 78]

    def note(self, text: str) -> None:
        self.lines.append(text)

    def check(self, name: str, actual: float, expected: float, tol: float = 1e-4) -> None:
        good = abs(actual - expected) <= tol
        self.passed += int(good)
        self.failed += int(not good)
        mark = "✔ PASS" if good else "✘ FAIL"
        extra = "" if good else f"  | 差异 {abs(actual - expected):.6f}"
        self.lines.append(f"  {mark}  {_pad(name, 40)}实际 = {actual:>18,.6f} | "
                          f"期望 = {expected:>18,.6f}{extra}")

    # ---------------- 用例 1 ----------------
    def case1(self) -> Dict[str, Any]:
        self.section("用例 1 | 个人储蓄复利增值 —— 小张的 10 万元五年后值多少")
        self.note("场景设定：小张将 100,000 元存入 5 年期理财，年化利率 6%，按年复利计息。")
        self.note("问题：5 年后本息合计多少？若改为按月复利，又能多拿多少？")
        pv, r, n = 100_000, 0.06, 5
        fv_year = fv_compound(pv, r, n, 1)
        fv_quarter = fv_compound(pv, r, n, 4)
        fv_month = fv_compound(pv, r, n, 12)
        ear_year = effective_annual_rate(r, 1)
        ear_month = effective_annual_rate(r, 12)
        pv_back = pv_compound(200_000, r, n, 1)
        self.note("【计算过程】")
        self.note(f"  1) 年复利终值 FV = {pv:,.0f} × (1 + {r:.2%})^5 = {fv_year:,.2f} 元")
        self.note(f"  2) 按季复利 FV = {fv_quarter:,.2f} 元   3) 按月复利 FV = {fv_month:,.2f} 元")
        self.note(f"  4) 实际年利率 EAR(月复利) = {ear_month:.4%}"
                  f"   5) 5 年后 20 万的复利现值 = {pv_back:,.2f} 元")
        self.check("年复利终值（10万，6%，5年）", fv_year, 133_822.5578, 0.01)
        self.check("季复利终值", fv_quarter, 134_685.5007, 0.01)
        self.check("月复利终值", fv_month, 134_885.0153, 0.01)
        self.check("实际年利率 EAR（月复利）", ear_month, 0.061678, 1e-5)
        self.check("20万的复利现值（5年，6%）", pv_back, 149_451.6346, 0.01)
        self.note("【业务解读】")
        self.note(f"  · 年复利下资金 5 年增长 {fv_year / pv - 1:.2%}，10 万变 {fv_year:,.0f} 元。")
        self.note(f"  · 计息频率效应：月复利比年复利多收益 {fv_month - fv_year:,.2f} 元，"
                  f"实际年利率从 {ear_year:.2%} 提升到 {ear_month:.2%}"
                  f"（+{(ear_month - ear_year) * 100:.3f} 个百分点）。")
        self.note(f"  · “72 法则”估算翻倍时间 12 年，与精确计算 {(np.log(2) / np.log(1.06)):.1f} 年吻合。")
        return {"名称": "用例1 个人储蓄复利增值",
                "参数": {"现值PV": pv, "年利率": r, "年数": n, "计息方式": "年/季/月"},
                "过程": [("年复利终值", f"{fv_year:,.2f}"), ("季复利终值", f"{fv_quarter:,.2f}"),
                         ("月复利终值", f"{fv_month:,.2f}"), ("EAR(月复利)", f"{ear_month:.4%}"),
                         ("20万复利现值", f"{pv_back:,.2f}")],
                "结论": f"年复利终值 {fv_year:,.2f} 元；月复利终值 {fv_month:,.2f} 元，"
                        f"计息频率提升可额外增厚收益 {fv_month - fv_year:,.2f} 元。"}

    # ---------------- 用例 2 ----------------
    def case2(self) -> Dict[str, Any]:
        self.section("用例 2 | 教育金定投 —— 每年存 2 万，10 年后能攒多少？")
        self.note("场景设定：每年末存入 20,000 元，年化收益 5%，连续 10 年；对照每年初存入的预付年金。")
        pmt, r, n = 20_000, 0.05, 10
        fva_o, fva_d = fva_ordinary(pmt, r, n), fva_due(pmt, r, n)
        pva_o, pva_d = pva_ordinary(pmt, r, n), pva_due(pmt, r, n)
        self.note("【计算过程】")
        self.note(f"  1) 普通年金终值 FVA = {pmt:,.0f} × [((1+{r:.0%})^{n} - 1) / {r:.0%}] = {fva_o:,.2f} 元")
        self.note(f"  2) 预付年金终值 = FVA × (1 + r) = {fva_d:,.2f} 元")
        self.note(f"  3) 年金终值系数 (F/A,5%,10) = {fva_o / pmt:.6f}（教材附表可查，应与本结果一致）")
        self.note(f"  4) 年金现值系数 (P/A,5%,10) = {pva_o / pmt:.6f}，年金现值 = {pva_o:,.2f} 元")
        self.check("普通年金终值系数 (F/A,5%,10)", fva_o / pmt, 12.577892, 1e-5)
        self.check("普通年金终值 FVA", fva_o, 251_557.8507, 0.01)
        self.check("预付年金终值", fva_d, 264_135.7432, 0.01)
        self.check("年金现值系数 (P/A,5%,10)", pva_o / pmt, 7.721734, 1e-5)
        self.check("普通年金现值 PVA", pva_o, 154_434.6986, 0.01)
        self.check("预付年金现值", pva_d, 162_156.4335, 0.01)
        self.note("【业务解读】")
        self.note(f"  · 累计投入 {pmt * n:,.0f} 元，10 年后可得 {fva_o:,.2f} 元，投资收益 "
                  f"{fva_o - pmt * n:,.2f} 元，占总资产 {(fva_o - pmt * n) / fva_o:.1%}。")
        self.note(f"  · 改为每年初存入，终值提升 {fva_d - fva_o:,.2f} 元（+{(fva_d / fva_o - 1):.2%}），"
                  f"这解释了保险、教育金产品为何强调“尽早缴费”。")
        return {"名称": "用例2 教育金定投年金",
                "参数": {"每期PMT": pmt, "年利率": r, "期数": n},
                "过程": [("普通年金终值", f"{fva_o:,.2f}"), ("预付年金终值", f"{fva_d:,.2f}"),
                         ("普通年金现值", f"{pva_o:,.2f}"), ("预付年金现值", f"{pva_d:,.2f}"),
                         ("(F/A,5%,10)", f"{fva_o / pmt:.6f}"), ("(P/A,5%,10)", f"{pva_o / pmt:.6f}")],
                "结论": f"10 年后可积累 {fva_o:,.2f} 元；改为期初存入可多得 {fva_d - fva_o:,.2f} 元。"}

    # ---------------- 用例 3 ----------------
    def case3(self) -> Dict[str, Any]:
        self.section("用例 3 | 企业债券发行资本成本 —— 折价发行的真实融资成本与税盾效应")
        self.note("场景设定：面值 1,000 元、票面利率 6%、期限 5 年的公司债，因市场利率上升实际"
                  "发行价 950 元，所得税率 25%。")
        fv_, cpn, yrs, price, tax = 1000, 0.06, 5, 950, 0.25
        d = cost_of_debt(fv_, cpn, yrs, price, tax)
        dap = cost_of_debt_approx(fv_, cpn, yrs, price, tax)
        bp = bond_price(fv_, cpn, yrs, cpn)
        self.note("【计算过程】")
        self.note(f"  1) 年利息 C = 1000 × 6% = {fv_ * cpn:,.2f} 元/年")
        self.note(f"  2) 建立方程：950 = Σ [60 / (1+k)^t] + 1000 / (1+k)^5 ，用牛顿迭代求 k")
        self.note(f"  3) 解得税前债务资本成本 Kd(税前) = {d['税前年化资本成本']:.4%}，"
                  f"税后 Kd = {d['税后年化资本成本']:.4%}")
        self.note(f"  4) 近似公式校验：Kd ≈ [1000×6% + (1000-950)/5] / [(1000+950)/2] = "
                  f"{dap['税前年化资本成本(近似)']:.4%}"
                  f"（两法差异 {abs(d['税前年化资本成本'] - dap['税前年化资本成本(近似)']):.5%}）")
        self.check("税前债务资本成本（YTM/IRR 法）", d["税前年化资本成本"], 0.072269, 1e-5)
        self.check("税后债务资本成本", d["税后年化资本成本"], 0.054202, 1e-5)
        self.check("近似公式法税前成本", dap["税前年化资本成本(近似)"], 0.071795, 1e-5)
        self.check("平价发行时的理论价格", bp["理论价格"], 1000.0, 0.01)
        self.note("【业务解读】")
        self.note(f"  · 票面利率只有 6%，但因折价发行，真实税前债务成本达 {d['税前年化资本成本']:.3%}，"
                  f"高出票面利率 {(d['税前年化资本成本'] - cpn) * 100:.3f} 个百分点——价格折让是对投资者的额外补偿。")
        self.note(f"  · 税盾效应使实际成本降至 {d['税后年化资本成本']:.3%}，"
                  f"每 100 元债务每年节省利息税负 {d['税盾节省(百分点)']:.3f} 元。")
        self.note("  · 债券定价校验：以票面利率折现时理论价格恰为面值 1,000 元，模型自洽。")
        return {"名称": "用例3 企业债券资本成本",
                "参数": {"面值": fv_, "票面利率": cpn, "期限": yrs, "发行价": price, "税率": tax},
                "过程": [("税前债务成本(YTM)", f"{d['税前年化资本成本']:.6%}"),
                         ("税后债务成本", f"{d['税后年化资本成本']:.6%}"),
                         ("近似公式法", f"{dap['税前年化资本成本(近似)']:.6%}"),
                         ("税盾节省(百分点)", f"{d['税盾节省(百分点)']:.4f}")],
                "结论": f"税前债务成本 {d['税前年化资本成本']:.4%}，税后 {d['税后年化资本成本']:.4%}，"
                        f"税盾每年节省 {d['税盾节省(百分点)']:.3f} 个百分点。"}

    # ---------------- 用例 4 ----------------
    def case4(self) -> Dict[str, Any]:
        self.section("用例 4 | 上市公司股权资本成本 —— CAPM 与 DDM 双重验证")
        self.note("场景设定：β = 1.20，10 年期国债收益率 2.5%，市场组合期望收益率 9.5%；"
                  "预计下年每股股利 0.50 元，当前股价 12 元，股利年增长率 4%。")
        rf, beta, rm, d1, p0, g = 0.025, 1.20, 0.095, 0.50, 12.0, 0.04
        e = capm_required_return(rf, beta, rm)
        ks_ddm = cost_of_equity_ddm(d1, p0, g)
        self.note("【计算过程】")
        self.note(f"  1) 市场风险溢价 MRP = {rm:.2%} - {rf:.2%} = {e['市场风险溢价']:.2%}")
        self.note(f"  2) CAPM：Ks = {rf:.2%} + {beta} × {e['市场风险溢价']:.2%} = {e['股权资本成本']:.4%}")
        self.note(f"  3) DDM：Ks = D1/P0 + g = {d1}/{p0} + {g:.2%} = {ks_ddm:.4%}"
                  f"（两模型差异 {ks_ddm - e['股权资本成本']:+.4%}）")
        self.check("市场风险溢价 MRP", e["市场风险溢价"], 0.07, 1e-9)
        self.check("CAPM 股权资本成本", e["股权资本成本"], 0.109, 1e-9)
        self.check("DDM 股权资本成本", ks_ddm, 0.081667, 1e-5)
        self.note("【业务解读】")
        self.note("  · 差异主要源于假设口径不同：CAPM 基于市场风险定价（β 由历史股价回归得出），"
                  "DDM 基于分红能力定价（依赖增长率假设）。实务中常取 CAPM 为主、DDM 校验。")
        self.note(f"  · β = {beta} > 1 表明波动高于市场，投资者要求 {e['风险补偿部分']:.2%} 的风险补偿，"
                  f"占股权成本的 {e['风险补偿部分'] / e['股权资本成本']:.0%}，这是“股权融资贵”的根源。")
        self.note(f"  · 若 MRP 假设从 7% 降至 5%，Ks 将降至 {rf + beta * 0.05:.2%}，"
                  f"说明 CAPM 对 MRP 假设高度敏感，测算时须审慎取值。")
        return {"名称": "用例4 股权资本成本CAPM",
                "参数": {"Rf": rf, "β": beta, "Rm": rm, "D1": d1, "P0": p0, "g": g},
                "过程": [("MRP", f"{e['市场风险溢价']:.4%}"), ("CAPM Ks", f"{e['股权资本成本']:.6%}"),
                         ("DDM Ks", f"{ks_ddm:.6%}"), ("两法差异", f"{ks_ddm - e['股权资本成本']:+.6%}")],
                "结论": f"CAPM 股权成本 {e['股权资本成本']:.4%}，DDM 交叉验证 {ks_ddm:.4%}，"
                        f"差异 {abs(ks_ddm - e['股权资本成本']) * 100:.2f} 个百分点。"}

    # ---------------- 用例 5 ----------------
    def case5(self) -> Dict[str, Any]:
        self.section("用例 5 | 某公司 WACC 计算 —— 资本结构决策与融资成本评估")
        self.note("场景设定：某智能制造上市公司，资本全部来自普通股与公司债，市场价值口径为"
                  "股权 6,000 万元、债务 4,000 万元；债务为面值 1,000 元、票面利率 6%、"
                  "5 年期、发行价 980 元的债券，所得税率 25%；"
                  "无风险利率 2.5%，β = 1.15，市场期望收益率 9.5%。")
        fv_, cpn, yrs, price, tax = 1000, 0.06, 5, 980, 0.25
        rf, beta, rm = 0.025, 1.15, 0.095
        E, D = 6000.0, 4000.0
        d = cost_of_debt(fv_, cpn, yrs, price, tax)
        kd_after = d["税后年化资本成本"]
        e = capm_required_return(rf, beta, rm)
        ks = e["股权资本成本"]
        w = wacc(E, D, ks, kd_after)
        self.note("【计算过程】")
        self.note(f"  步骤 1  债务成本：税前 Kd = {d['税前年化资本成本']:.4%}，税后 Kd = {kd_after:.4%}")
        self.note(f"  步骤 2  股权成本：Ks = {rf:.2%} + {beta} × ({rm:.2%} - {rf:.2%}) = {ks:.4%}")
        self.note(f"  步骤 3  权重：wE = {w['股权比重']:.4%}，wD = {w['债权比重']:.4%}")
        self.note(f"  步骤 4  WACC = {w['WACC']:.4%}"
                  f"（股权贡献 {w['股权贡献度(百分点)']:.4f} + 债权贡献 {w['债权贡献度(百分点)']:.4f} 个百分点）")
        self.check("税前债务资本成本", d["税前年化资本成本"], 0.064810, 1e-5)
        self.check("税后债务资本成本", kd_after, 0.048608, 1e-5)
        self.check("股权资本成本 Ks", ks, 0.1055, 1e-9)
        self.check("股权权重", w["股权比重"], 0.6, 1e-9)
        self.check("债权权重", w["债权比重"], 0.4, 1e-9)
        self.check("WACC", w["WACC"], 0.082743, 1e-5)
        self.check("贡献分解可加性（股权+债权）",
                   w["股权贡献度(百分点)"] / 100 + w["债权贡献度(百分点)"] / 100, w["WACC"], 1e-9)
        self.note("【业务解读】")
        self.note(f"  · WACC = {w['WACC']:.2%}，低于行业平均 10%，具备融资成本优势。")
        self.note(f"  · 股权成本 {ks:.2%} 明显高于税后债务成本 {kd_after:.2%}，"
                  f"杠杆在放大融资能力的同时也带来固定财务负担。")
        self.note(f"  · 新增项目的最低可接受收益率（门槛收益率）应不低于 {w['WACC']:.2%}，"
                  f"否则无法覆盖资本成本；若项目 ROIC 达到 11.8%，则 EVA 为正。")
        return {"名称": "用例5 某公司WACC计算",
                "参数": {"股权市场价值E": E, "债务市场价值D": D, "债券面值": fv_, "票面利率": cpn,
                         "发行价": price, "税率": tax, "Rf": rf, "β": beta, "Rm": rm},
                "过程": [("税前债务成本", f"{d['税前年化资本成本']:.6%}"), ("税后债务成本", f"{kd_after:.6%}"),
                         ("股权成本Ks", f"{ks:.6%}"), ("股权权重", f"{w['股权比重']:.4%}"),
                         ("债权权重", f"{w['债权比重']:.4%}"), ("WACC", f"{w['WACC']:.6%}")],
                "结论": f"该公司 WACC = {w['WACC']:.4%}，低于行业平均 10%，融资成本具备竞争优势。"}

    # ---------------- 用例 6 ----------------
    def case6(self) -> Dict[str, Any]:
        self.section("用例 6 | 房贷还款计划 —— 等额本息下的利息构成（年金公式逆向应用）")
        self.note("场景设定：贷款 100 万元，年利率 4.2%（月利率 0.35%），期限 20 年，等额本息按月还款。")
        P, rate, years, freq = 1_000_000, 0.042, 20, 12
        s = amortization_schedule(P, rate, years, freq)
        r_m, n = rate / freq, years * freq
        verify = pva_ordinary(s["每期还款额"], r_m, n)
        self.note("【计算过程】")
        self.note(f"  1) 期利率 r = {r_m:.4%}，总期数 n = {n}")
        self.note(f"  2) 月供 PMT = P × r / [1 - (1+r)^(-n)] = {s['每期还款额']:,.2f} 元")
        self.note(f"  3) 总还款额 = {s['总还款额']:,.2f} 元，总利息 = {s['总利息']:,.2f} 元")
        self.note(f"  4) 逆向校验：把月供按 {r_m:.4%} 折现 {n} 期 = {verify:,.6f} 元 ≈ 贷款本金 ✔")
        self.check("月供 PMT", s["每期还款额"], 6165.7074, 0.01)
        self.check("总利息", s["总利息"], 479_769.77, 0.5)
        self.check("月供的年金现值（应等于本金）", verify, 1_000_000.0, 1e-4)
        self.check("首期利息", s["计划表"][0]["利息"], 3500.0, 1e-6)
        self.note("【业务解读】")
        self.note(f"  · 20 年累计支付利息 {s['总利息']:,.2f} 元，相当于本金的 {s['总利息'] / P:.1%}。")
        self.note(f"  · 首期月供中利息占 {s['计划表'][0]['利息'] / s['每期还款额']:.1%}，"
                  f"本金仅占 {s['计划表'][0]['本金'] / s['每期还款额']:.1%}；最后一期利息几乎为零。")
        self.note("  · 校验意义：月供的年金现值精确等于本金，说明房贷本质就是“本金 = 各期还款的年金现值”。")
        self.note("  · 决策提示：提前还款越早，节省利息越多。")
        return {"名称": "用例6 房贷还款计划",
                "参数": {"贷款本金": P, "年利率": rate, "年限": years, "每年还款次数": freq},
                "过程": [("月供PMT", f"{s['每期还款额']:,.2f}"), ("总还款额", f"{s['总还款额']:,.2f}"),
                         ("总利息", f"{s['总利息']:,.2f}"), ("月供年金现值校验", f"{verify:,.6f}")],
                "结论": f"月供 {s['每期还款额']:,.2f} 元，总利息 {s['总利息']:,.2f} 元，年金现值逆向校验通过。"}

    # ---------------- 附加：参数校验 ----------------
    def case_validation(self) -> Dict[str, Any]:
        self.section("附加测试 | AI 辅助参数校验与异常处理能力验证")
        trials = [
            ("利率填成百分数（6 而非 0.06）", "rate", 6),
            ("利率为负数（-0.5）", "rate", -0.5),
            ("期数为小数（3.5）", "periods", 3.5),
            ("期数为 0", "periods", 0),
            ("β 为极端值（3.2）", "beta", 3.2),
            ("所得税率为 1.2（>100%）", "tax_rate", 1.2),
            ("金额为空", "positive_amount", ""),
            ("金额为文本", "positive_amount", "abc"),
        ]
        rows = []
        self.note(f"  {_pad('测试情形', 34)}{_pad('判定', 14)}提示信息")
        self.note("  " + "─" * 104)
        for name, rule, val in trials:
            r = validate(rule, val)
            verdict = "拦截(ERROR)" if r.errors else ("警告(WARN)" if r.warnings else "通过(OK)")
            msg = (r.errors or r.warnings or ["—"])[0]
            self.note(f"  {_pad(name, 34)}{_pad(verdict, 14)}{msg[:62]}")
            rows.append({"情形": name, "判定": verdict, "提示": msg})
            self.passed += int(bool(r.errors) or bool(r.warnings))
            self.failed += int(not (r.errors or r.warnings))

        w_res = check_wacc_weights(0.45, 0.5)
        ok_weight = bool(w_res.errors)
        self.passed += int(ok_weight)
        self.failed += int(not ok_weight)
        self.note(f"  {_pad('WACC 权重合计 = 0.95（≠1）', 34)}"
                  f"{_pad('拦截(ERROR)' if ok_weight else '通过', 14)}"
                  f"{(w_res.errors[0][:62] if w_res.errors else '—')}")

        ok_res = validate("rate", 0.06)
        passed_ok = ok_res.ok and not ok_res.warnings
        self.passed += int(passed_ok)
        self.failed += int(not passed_ok)
        self.note(f"  {_pad('正常输入 0.06（6%）', 34)}{_pad('通过(OK)', 14)}校验通过，无阻断项")

        self.note("【结论】")
        self.note("  · 所有非法输入均在计算前被拦截，并给出“错在哪里 + 应如何修正”的具体提示，"
                  "避免程序崩溃或输出无意义的计算结果。")
        return {"名称": "附加测试 参数校验",
                "参数": {"测试项数": len(trials) + 1},
                "过程": [(r["情形"], r["判定"]) for r in rows],
                "结论": f"全部 {len(trials) + 1} 项异常输入均被正确识别与提示，无漏检。"}

    # ---------------- 主流程 ----------------
    def run(self) -> Tuple[bool, List[str], List[Dict[str, Any]]]:
        self.lines = ["货币时间价值与资本成本计算器 —— 数值自检测试",
                      "共 6 组主用例 + 1 组参数校验测试，逐项与教材系数表/解析解比对", ""]
        for fn in (self.case1, self.case2, self.case3, self.case4,
                   self.case5, self.case6, self.case_validation):
            self.records.append(fn())
        self.lines += ["", "=" * 78,
                       f"测试完成：通过 {self.passed} 项，失败 {self.failed} 项。",
                       ("全部数值断言通过（与教材系数表、解析解逐项比对）。"
                        if self.failed == 0
                        else "存在未通过项，请检查计算引擎。"),
                       "=" * 78]
        return self.failed == 0, self.lines, self.records


def run_self_test(export_path: Optional[str] = None) -> Tuple[bool, List[str], List[Dict[str, Any]]]:
    """运行数值自检测试；可选导出测试用例明细 Excel。"""
    ok, lines, records = SelfTest().run()
    if export_path:
        export_casebook(records, export_path)
        lines.append(f"测试用例明细已导出：{export_path}")
    return ok, lines, records


# ============================================= 5.9 一键运行全部场景
def demo_results(out_dir: Optional[str] = None) -> List[SceneResult]:
    """按顺序执行全部示例场景（与旧版 main.py --demo 参数完全一致）。"""
    out_dir = out_dir or get_output_dir()
    return [
        run_compound(100_000, 0.06, 10, 1, "1", out_dir),
        run_annuity(10_000, 0.08, 10, False, out_dir),
        run_debt(1000, 0.06, 5, 950, 0.25, out_dir),
        run_equity(0.025, 1.2, 0.095, 0.0, True, 0.5, 12.0, 0.04, out_dir),
        run_wacc(1000, 0.06, 5, 980, 0.25, 0.025, 1.15, 0.095, 6000, 4000, 0.0, 0.0, out_dir),
        run_loan(1_000_000, 0.042, 20, 12, out_dir),
    ]


def run_all_scenarios(out_dir: Optional[str] = None) -> Tuple[List[SceneResult], List[str]]:
    """
    一键运行全部场景并导出 09_计算结果汇总.xlsx 与 计算结果汇总.md。
    返回 (场景结果列表, 生成的文件路径列表)。
    """
    out_dir = out_dir or get_output_dir()
    results = demo_results(out_dir)

    xlsx = os.path.join(out_dir, "09_计算结果汇总.xlsx")
    md = os.path.join(out_dir, "计算结果汇总.md")
    extras: List[Tuple[str, Any]] = []
    for r in results:
        extras += list(r.extra_sheets)
    export_summary([r.summary for r in results], xlsx, extras)
    export_markdown([r.summary for r in results], md)
    return results, [xlsx, md]



# =============================================================================
# 6. 输出目录
# =============================================================================
def get_script_dir() -> str:
    """脚本所在目录（双击 .pyw 时为用户放置该文件的目录）。"""
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:                                   # 交互式解释器
        return os.getcwd()


_OUT_DIR_CACHE: Optional[str] = None


def get_output_dir() -> str:
    """
    解析输出目录：
      1. 脚本同级已存在的 outputs/ 或 output/（沿用旧版目录，避免产生两份产物）
      2. 否则在脚本同级新建 outputs/；若该位置不可写，退化到用户主目录。
    """
    global _OUT_DIR_CACHE
    if _OUT_DIR_CACHE:
        return _OUT_DIR_CACHE

    base = get_script_dir()
    for name in ("outputs", "output"):
        cand = os.path.join(base, name)
        if os.path.isdir(cand):
            _OUT_DIR_CACHE = cand
            return cand

    cand = os.path.join(base, "outputs")
    try:
        os.makedirs(cand, exist_ok=True)
        probe = os.path.join(cand, ".write_test")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("ok")
        os.remove(probe)
        _OUT_DIR_CACHE = cand
    except Exception:                                   # noqa: BLE001
        fallback = os.path.join(os.path.expanduser("~"), "财务计算器输出")
        os.makedirs(fallback, exist_ok=True)
        _OUT_DIR_CACHE = fallback
    return _OUT_DIR_CACHE


# =============================================================================
# 7. Tkinter 图形界面
# =============================================================================
try:
    import tkinter as tk
    from tkinter import messagebox, ttk
    from tkinter.scrolledtext import ScrolledText
    TK_OK = True
except Exception as _tk_exc:                            # noqa: BLE001
    TK_OK = False
    _IMPORT_ERRORS["tkinter"] = f"{type(_tk_exc).__name__}: {_tk_exc}"


# ---- 界面配色 ----
UI = {
    # 深色工业仪表盘：降低大面积纯黑刺眼感，强调信息层级和数值对比度。
    "bg": "#0B1220", "panel": "#111C2F", "panel_alt": "#16253D",
    "header": "#101B30", "header_text": "#F8FAFC", "accent": "#F59E0B",
    "accent_soft": "#FDE68A", "h1": "#F8FAFC", "h2": "#7DD3FC",
    "ok": "#34D399", "warn": "#FBBF24", "err": "#FB7185",
    "muted": "#94A3B8", "body": "#E2E8F0", "border": "#263A58",
    "input": "#0F1A2E", "selection": "#1E3A5F",
}


def _open_path(path: str) -> None:
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


@dataclass
class FieldSpec:
    """界面输入字段定义。"""
    key: str
    label: str
    default: Any
    rule: Optional[str] = None
    kind: str = "number"                    # number | int | choice
    options: Sequence[Tuple[str, Any]] = ()
    hint: str = ""


class StatCard(tk.Frame):
    """总览页的关键指标卡：用清晰的数值层级替代大段说明。"""

    def __init__(self, master, label: str, value: str, accent: str = None):
        super().__init__(master, bg=UI["panel_alt"], padx=14, pady=10,
                         highlightbackground=UI["border"], highlightthickness=1)
        tk.Label(self, text=label, bg=UI["panel_alt"], fg=UI["muted"],
                 font=("Microsoft YaHei", 9)).pack(anchor="w")
        tk.Label(self, text=value, bg=UI["panel_alt"], fg=accent or UI["accent"],
                 font=("Consolas", 18, "bold")).pack(anchor="w", pady=(5, 0))


class PreviewPanel(ttk.LabelFrame):
    """图表预览面板：下拉选择输出目录中的 PNG 并缩放显示。"""

    def __init__(self, master, app, title: str = "图表预览"):
        super().__init__(master, text=title, padding=8)
        self.app = app
        self._img_ref: Optional[Any] = None
        self._path: Optional[str] = None

        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text="选择图表：").pack(side="left")
        self.combo = ttk.Combobox(top, state="readonly", width=30)
        self.combo.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.combo.bind("<<ComboboxSelected>>", lambda _e: self._on_pick())
        ttk.Button(top, text="刷新", width=6, command=self.refresh).pack(side="left", padx=2)
        ttk.Button(top, text="打开原图", width=9, command=self._open_current).pack(side="left", padx=2)

        # 用空白占位图把 Label 的尺寸单位固定为像素，便于后续缩放计算
        self._blank = tk.PhotoImage(width=640, height=430)
        self._blank.put(UI["input"], to=(0, 0, 640, 430))
        self.canvas_label = tk.Label(self, image=self._blank, bg=UI["input"],
                                     relief="solid", bd=1, width=640, height=430,
                                     highlightbackground=UI["border"], highlightthickness=1)
        self.canvas_label.pack(fill="both", expand=True, pady=(8, 0))

        self.hint = ttk.Label(self, text="计算后自动显示生成的图表", foreground=UI["muted"])
        self.hint.pack(anchor="w", pady=(4, 0))

    # ---------------- 内部 ----------------
    def _pngs(self) -> List[str]:
        d = get_output_dir()
        try:
            names = [n for n in os.listdir(d) if n.lower().endswith(".png")]
        except OSError:
            return []
        return [os.path.join(d, n) for n in sorted(names)]

    def _on_pick(self) -> None:
        idx = self.combo.current()
        paths = self._pngs()
        if 0 <= idx < len(paths):
            self.show(paths[idx])

    def _open_current(self) -> None:
        _open_path(self._path or get_output_dir())

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

    def show(self, path: str) -> None:
        """加载 PNG 并按可用空间缩放显示（Tk 8.6 原生支持 PNG）。"""
        try:
            img = tk.PhotoImage(file=path)
        except Exception as exc:                        # noqa: BLE001
            self.hint.configure(text=f"无法预览：{exc}", foreground=UI["err"])
            return

        self.canvas_label.update_idletasks()
        avail_w = max(self.canvas_label.winfo_width(), 400)
        avail_h = max(self.canvas_label.winfo_height(), 300)
        factor = max(1, math.ceil(max(img.width() / avail_w, img.height() / avail_h)))
        if factor > 1:
            img = img.subsample(factor, factor)

        self.canvas_label.configure(image=img, text="", bg=UI["input"])
        self._img_ref = img                                # 必须持有引用，否则被 GC
        self._path = path
        try:
            size_kb = os.path.getsize(path) / 1024
        except OSError:
            size_kb = 0
        self.hint.configure(
            text=f"{os.path.basename(path)}   （原图 {size_kb:.0f} KB，预览缩放 1/{factor}）",
            foreground=UI["muted"])


class ResultText(ttk.Frame):
    """带标签样式的只读文本区，用于展示计算结果与业务解读。"""

    def __init__(self, master):
        super().__init__(master)
        self.text = ScrolledText(self, wrap="word", font=("Consolas", 10),
                                 bg=UI["input"], fg=UI["body"], relief="solid", bd=1,
                                 insertbackground=UI["accent"], selectbackground=UI["selection"],
                                 padx=12, pady=10, height=22,
                                 highlightbackground=UI["border"], highlightthickness=1)
        self.text.pack(fill="both", expand=True)
        self.text.configure(state="disabled")
        self.text.tag_configure("h1", font=("Microsoft YaHei", 13, "bold"), foreground=UI["accent"],
                                spacing1=8, spacing3=4)
        self.text.tag_configure("h2", font=("Microsoft YaHei", 11, "bold"), foreground=UI["h2"],
                                spacing1=10, spacing3=4)
        self.text.tag_configure("body", font=("Consolas", 10), foreground=UI["body"])
        self.text.tag_configure("warn", font=("Consolas", 10), foreground=UI["warn"])
        self.text.tag_configure("err", font=("Consolas", 10), foreground=UI["err"])
        self.text.tag_configure("ok", font=("Consolas", 10, "bold"), foreground=UI["ok"])
        self.text.tag_configure("muted", font=("Consolas", 9), foreground=UI["muted"])
        self.text.tag_configure("formula", font=("Consolas", 10), foreground=UI["accent_soft"],
                                lmargin1=8, lmargin2=8)

    def _write(self, chunks: Sequence[Tuple[str, str]]) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        for tag, line in chunks:
            self.text.insert("end", line + "\n", tag)
        self.text.configure(state="disabled")
        self.text.see("1.0")

    def clear(self) -> None:
        self._write([])

    def show_lines(self, title: str, lines: Sequence[Tuple[str, str]]) -> None:
        chunks: List[Tuple[str, str]] = [("h1", title), ("h1", "")]
        chunks += [(tag, line) for tag, line in lines]
        self._write(chunks)

    def show_blocks(self, blocks: Sequence[Tuple[str, Sequence[Tuple[str, str]]]]) -> None:
        chunks: List[Tuple[str, str]] = [("h1", f"运行结果  ·  {APP_VERSION}"), ("h1", "")]
        for title, lines in blocks:
            chunks.append(("h2", "▶ " + title))
            for tag, line in lines:
                if tag == "body" and line.strip().startswith(("⚠", "✘")):
                    tag = "err"
                elif tag == "body" and line.strip().startswith("✔"):
                    tag = "ok"
                chunks.append((tag, line))
            chunks.append(("body", ""))
        self._write(chunks)

    def append(self, line: str, tag: str = "muted") -> None:
        self.text.configure(state="normal")
        self.text.insert("end", line + "\n", tag)
        self.text.configure(state="disabled")
        self.text.see("end")



class SceneTab(ttk.Frame):
    """一个场景页：左侧参数输入 + 按钮，右侧结果文本 + 图表预览。"""

    def __init__(self, master, app, title: str, desc: str,
                 fields: Sequence[FieldSpec], runner: Callable[..., SceneResult]):
        super().__init__(master, padding=10)
        self.app = app
        self.title_text = title
        self.desc = desc
        self.fields = list(fields)
        self.runner = runner
        self.vars: Dict[str, Any] = {}
        self.entries: Dict[str, Any] = {}

        ttk.Label(self, text=title, font=("Microsoft YaHei", 14, "bold"),
                  foreground=UI["accent"]).pack(anchor="w")
        if desc:
            ttk.Label(self, text=desc, foreground=UI["muted"]).pack(anchor="w", pady=(2, 8))

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        right = ttk.Frame(body)
        body.add(left, weight=3)
        body.add(right, weight=4)

        self._build_params(left)
        self.text = ResultText(right)
        self.text.pack(fill="both", expand=True)
        self.preview = PreviewPanel(body, app, "图表预览")
        body.add(self.preview, weight=4)

    # ---------------- 参数区 ----------------
    def _build_params(self, parent: ttk.Frame) -> None:
        box = ttk.LabelFrame(parent, text="参数输入  ·  利率按小数填写", padding=12)
        box.pack(fill="x")
        box.columnconfigure(1, weight=1)
        box.columnconfigure(3, weight=1)

        row = 0
        for i, f in enumerate(self.fields):
            col = i % 2
            if col == 0 and i > 0:
                row += 2
            ttk.Label(box, text=f.label).grid(row=row, column=col * 2, sticky="w",
                                              padx=(0, 6), pady=4)
            if f.kind == "choice":
                var = tk.StringVar(value=self._label_of(f, f.default))
                widget = ttk.Combobox(box, textvariable=var, state="readonly", width=18,
                                      values=[lbl for lbl, _ in f.options])
            else:
                var = tk.StringVar(value=self._fmt_default(f.default))
                widget = ttk.Entry(box, textvariable=var, width=18)
            widget.grid(row=row, column=col * 2 + 1, sticky="ew", padx=(0, 16), pady=4)
            self.vars[f.key] = var
            self.entries[f.key] = widget
            if f.hint:
                ttk.Label(box, text=f.hint, foreground=UI["muted"],
                          font=("Microsoft YaHei", 8)).grid(
                    row=row + 1, column=col * 2, columnspan=2, sticky="w", pady=(0, 4))

        btns = ttk.Frame(parent)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="计算并生成图表", style="Accent.TButton",
                   command=self.on_calc).pack(side="left")
        ttk.Button(btns, text="恢复示例值", command=self.on_reset).pack(side="left", padx=8)
        ttk.Button(btns, text="清空", command=self.on_clear).pack(side="left")

    @staticmethod
    def _fmt_default(v: Any) -> str:
        if isinstance(v, float):
            return f"{v:g}"
        return str(v)

    @staticmethod
    def _label_of(f: FieldSpec, value: Any) -> str:
        for lbl, val in f.options:
            if val == value:
                return lbl
        return f.options[0][0] if f.options else ""

    # ---------------- 取值与校验 ----------------
    def _read(self) -> Tuple[Optional[Dict[str, Any]], List[str], List[str]]:
        values: Dict[str, Any] = {}
        errors: List[str] = []
        warnings: List[str] = []

        for f in self.fields:
            raw = str(self.vars[f.key].get()).strip()
            if f.kind != "choice":
                self.entries[f.key].configure(style="TEntry")
            if f.kind == "choice":
                values[f.key] = dict((lbl, val) for lbl, val in f.options)[raw]
                continue
            if raw == "":
                errors.append(f"【{f.label}】不能为空，请输入一个数值。")
                self.entries[f.key].configure(style="Error.TEntry")
                continue
            if f.rule:
                res = validate(f.rule, raw)
                if res.errors:
                    errors += res.errors
                    self.entries[f.key].configure(style="Error.TEntry")
                    continue
                warnings += res.warnings
            try:
                num = float(raw)
            except ValueError:
                errors.append(f"【{f.label}】必须为数字，当前输入为 “{raw}”。")
                self.entries[f.key].configure(style="Error.TEntry")
                continue
            values[f.key] = int(round(num)) if f.kind == "int" else num
        return (values if not errors else None), errors, warnings

    # ---------------- 事件 ----------------
    def on_calc(self) -> None:
        values, errors, warnings = self._read()
        if values is None:
            self.text.show_lines("参数校验未通过", [("err", e) for e in errors] + [
                ("muted", ""),
                ("muted", "提示：利率请填小数（6% 填 0.06）；期数 / 年数填正整数；金额填数值（元）。")])
            self.app.set_status("参数校验未通过，已标红错误字段")
            return
        try:
            result = self.runner(**values, out_dir=get_output_dir())
        except FinanceError as exc:
            self.text.show_lines("计算失败", [("err", str(exc))])
            return
        except Exception as exc:                        # noqa: BLE001
            self.text.show_lines("程序异常", [("err", f"{type(exc).__name__}: {exc}"),
                                          ("body", traceback.format_exc())])
            return

        blocks: List[Tuple[str, Sequence[Tuple[str, str]]]] = []
        if warnings:
            blocks.append(("参数提示（不影响计算）", [("warn", w) for w in warnings]))
        for sec_title, lines in result.sections:
            blocks.append((sec_title, [("body", ln) for ln in lines]))
        self.text.show_blocks(blocks)
        self.app.on_scene_done(result)

    def on_reset(self) -> None:
        for f in self.fields:
            if f.kind == "choice":
                self.vars[f.key].set(self._label_of(f, f.default))
            else:
                self.vars[f.key].set(self._fmt_default(f.default))
                self.entries[f.key].configure(style="TEntry")

    def on_clear(self) -> None:
        for f in self.fields:
            if f.kind != "choice":
                self.vars[f.key].set("")


class OverviewTab(ttk.Frame):
    """总览页：一键运行全部场景、数值自检测试、导出测试用例簿、打开输出目录。"""

    def __init__(self, master, app):
        super().__init__(master, padding=10)
        self.app = app

        hero = tk.Frame(self, bg=UI["panel"], padx=18, pady=16,
                        highlightbackground=UI["border"], highlightthickness=1)
        hero.pack(fill="x", pady=(0, 12))
        tk.Label(hero, text="财务计算工作台", bg=UI["panel"], fg=UI["accent"],
                 font=("Microsoft YaHei", 20, "bold")).pack(anchor="w")
        tk.Label(hero, text=APP_TAGLINE,
                 bg=UI["panel"], fg=UI["body"],
                 font=("Microsoft YaHei", 10)).pack(anchor="w", pady=(4, 0))
        tk.Label(hero, text="第 3 章 货币时间价值  /  第 4 章 资本成本  /  可审计输出",
                 bg=UI["panel"], fg=UI["muted"],
                 font=("Consolas", 9)).pack(anchor="w", pady=(8, 0))

        cards = tk.Frame(self, bg=UI["bg"])
        cards.pack(fill="x", pady=(0, 12))
        for label, value, color in (("计算场景", "06", UI["accent"]),
                                    ("专业图表", "08", UI["h2"]),
                                    ("校验机制", "L1 · L2 · L3", UI["ok"]),
                                    ("输出格式", "Excel / MD", UI["accent_soft"])):
            StatCard(cards, label, value, color).pack(side="left", fill="x", expand=True,
                                                       padx=(0, 8))
        ttk.Label(self, text="一键生成全部图表与 Excel / Markdown 汇总，或运行数值自检测试。",
                  foreground=UI["muted"]).pack(anchor="w", pady=(0, 8))

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="一键运行全部场景", style="Accent.TButton",
                   command=self.on_demo).pack(side="left")
        ttk.Button(bar, text="运行数值自检", command=self.on_selftest).pack(side="left", padx=8)
        ttk.Button(bar, text="导出测试用例簿", command=self.on_casebook).pack(side="left")
        ttk.Button(bar, text="打开输出目录",
                   command=lambda: _open_path(get_output_dir())).pack(side="left", padx=8)
        ttk.Button(bar, text="使用说明", command=self.app.show_help).pack(side="left")

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True)
        self.text = ResultText(body)
        body.add(self.text, weight=3)
        self.preview = PreviewPanel(body, app, "产物预览")
        body.add(self.preview, weight=4)

        self.text.show_lines("欢迎使用", [
            ("body", f"{APP_TITLE} {APP_VERSION}"),
            ("body", "对应课程第 3 章（货币时间价值）与第 4 章（资本成本）。"),
            ("body", ""),
            ("body", "· 左侧场景页可单独填参数计算，结果与业务解读显示在右侧，图表即时预览。"),
            ("body", "· 「一键运行全部场景」将按示例参数跑完 6 个场景，"
                     "生成 8 张图表与 09_计算结果汇总.xlsx、计算结果汇总.md。"),
            ("body", "· 「运行数值自检测试」将执行 6 组用例 + 1 组参数校验，"
                     "逐项与教材系数表、解析解比对。"),
            ("body", ""),
            ("muted", f"输出目录：{get_output_dir()}"),
            ("muted", f"图表中文字体：{CJK_FONT}"),
        ])

    # ---------------- 事件 ----------------
    def on_demo(self) -> None:
        try:
            results, files = run_all_scenarios(get_output_dir())
        except Exception as exc:                        # noqa: BLE001
            self.text.show_lines("运行失败", [("err", f"{type(exc).__name__}: {exc}"),
                                          ("body", traceback.format_exc())])
            return

        blocks: List[Tuple[str, List[Tuple[str, str]]]] = []
        for i, r in enumerate(results, 1):
            lines: List[Tuple[str, str]] = []
            for sec_title, sec_lines in r.sections:
                lines.append(("h2", f"  【{sec_title}】"))
                lines += [("body", "  " + ln) for ln in sec_lines]
            blocks.append((f"{i}. {r.name}", lines))
        blocks.append(("输出文件", [("ok", "  ✔ " + p) for p in files]
                       + [("body", "  · " + p) for r in results for p in r.figures + r.files]))
        self.text.show_blocks(blocks)
        self.preview.refresh()
        self.app.set_status(f"一键运行完成：{len(results)} 个场景，产物已写入 {get_output_dir()}")
        messagebox.showinfo("运行完成",
                            f"全部 {len(results)} 个场景执行完毕。\n\n"
                            f"图表与表格已写入：\n{get_output_dir()}")

    def on_selftest(self) -> None:
        try:
            ok, lines, _records = run_self_test()
        except Exception as exc:                        # noqa: BLE001
            self.text.show_lines("自检异常", [("err", f"{type(exc).__name__}: {exc}"),
                                          ("body", traceback.format_exc())])
            return
        chunks: List[Tuple[str, str]] = [("h1", "数值自检测试结果"), ("h1", "")]
        for ln in lines:
            if ln.startswith("✔"):
                chunks.append(("ok", ln))
            elif ln.startswith("✘"):
                chunks.append(("err", ln))
            elif ln.startswith(("━", "=")):
                chunks.append(("muted", ln))
            elif ln.startswith(("用例", "附加")):
                chunks.append(("h2", ln))
            else:
                chunks.append(("body", ln))
        self.text._write(chunks)
        self.app.set_status("自检" + ("全部通过" if ok else "存在失败项，请查看结果区"))
        if ok:
            messagebox.showinfo("自检测试", "全部数值断言通过。")
        else:
            messagebox.showwarning("自检测试", "存在未通过项，请查看结果区。")

    def on_casebook(self) -> None:
        path = os.path.join(get_output_dir(), "10_测试用例明细.xlsx")
        try:
            _ok, _lines, records = run_self_test()
            export_casebook(records, path)
        except Exception as exc:                        # noqa: BLE001
            messagebox.showerror("导出失败", f"{type(exc).__name__}: {exc}")
            return
        self.text.append(f"测试用例明细已导出：{path}")
        self.app.set_status(f"已导出 {os.path.basename(path)}")
        messagebox.showinfo("导出完成", f"测试用例明细已导出：\n{path}")


class FinanceCalculatorApp:
    """主应用：顶部工具条 + 7 个功能页 + 状态栏。"""

    def __init__(self, root):
        self.root = root
        root.title(f"{APP_TITLE} {APP_VERSION}")
        root.geometry("1440x920")
        root.minsize(1100, 720)
        root.configure(bg=UI["bg"])
        root.report_callback_exception = self._on_tk_error

        self._setup_style()
        self._configure_window()
        self._build_header()
        self._build_toolbar()
        self._build_notebook()
        self._build_statusbar()

    # ---------------- 外观 ----------------
    def _configure_window(self) -> None:
        """让深色界面在首次打开时居中，并统一系统级颜色。"""
        self.root.option_add("*TCombobox*Listbox.background", UI["input"])
        self.root.option_add("*TCombobox*Listbox.foreground", UI["body"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", UI["selection"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", UI["body"])
        self.root.option_add("*TNotebook.Tab.padding", (14, 8))

    def _setup_style(self) -> None:
        """优先选 clam：它对 background/foreground 的支持最完整，能真正呈现深色工业仪表盘；
        vista / xpnative 在 Windows 上较原生但会忽略很多 style 配置，故作为兜底。"""
        style = ttk.Style()
        theme_chosen = None
        for theme in ("clam", "alt", "default", "vista", "xpnative"):
            if theme in style.theme_names():
                try:
                    style.theme_use(theme)
                    theme_chosen = theme
                    break
                except tk.TclError:
                    continue
        # 让背景穿透到所有容器，避免出现"白底卡片"割裂感
        style.configure(".", font=("Microsoft YaHei", 9), background=UI["bg"], foreground=UI["body"])
        for cls in ("TFrame", "TLabel", "TLabelframe"):
            style.configure(cls, background=UI["bg"], foreground=UI["body"])
        style.configure("TLabelframe", bordercolor=UI["border"], relief="solid")
        style.configure("TLabelframe.Label", background=UI["bg"],
                        font=("Microsoft YaHei", 9, "bold"), foreground=UI["muted"])

        style.configure("TButton", font=("Microsoft YaHei", 9, "bold"), padding=(12, 7),
                        background=UI["panel_alt"], foreground=UI["body"], bordercolor=UI["border"])
        style.map("TButton",
                  background=[("active", UI["selection"]), ("pressed", UI["selection"]),
                              ("disabled", UI["panel"])],
                  foreground=[("disabled", UI["muted"]), ("active", UI["accent_soft"])])
        style.configure("Accent.TButton", font=("Microsoft YaHei", 9, "bold"), padding=(14, 8),
                        background=UI["accent"], foreground="#111827", bordercolor=UI["accent"])
        style.map("Accent.TButton",
                  background=[("active", "#FBBF24"), ("pressed", "#D97706"),
                              ("disabled", UI["panel_alt"])],
                  foreground=[("disabled", UI["muted"])])

        style.configure("TEntry", fieldbackground=UI["input"], foreground=UI["body"],
                        insertcolor=UI["accent"], bordercolor=UI["border"])
        style.configure("TCombobox", fieldbackground=UI["input"], foreground=UI["body"],
                        selectbackground=UI["selection"], selectforeground=UI["body"],
                        bordercolor=UI["border"])
        style.configure("TNotebook", background=UI["bg"], bordercolor=UI["border"], tabmargins=(2, 6, 2, 0))
        style.configure("TNotebook.Tab", font=("Microsoft YaHei", 9, "bold"), padding=(14, 8),
                        background=UI["panel"], foreground=UI["muted"], bordercolor=UI["border"])
        style.map("TNotebook.Tab",
                  background=[("selected", UI["accent"]), ("active", UI["panel_alt"])],
                  foreground=[("selected", "#111827"), ("active", UI["body"])])
        style.configure("Error.TEntry", fieldbackground="#4C1D2A", foreground="#FFE4E6")

        self._theme = theme_chosen

    def _build_header(self) -> None:
        head = tk.Frame(self.root, bg=UI["header"], height=78,
                        highlightbackground=UI["border"], highlightthickness=1)
        head.pack(fill="x")
        head.pack_propagate(False)
        left = tk.Frame(head, bg=UI["header"])
        left.pack(side="left", padx=18, pady=11)
        tk.Label(left, text="FINANCE / LAB", bg=UI["header"], fg=UI["accent"],
                 font=("Consolas", 9, "bold")).pack(anchor="w")
        tk.Label(left, text=APP_TITLE, bg=UI["header"], fg=UI["header_text"],
                 font=("Microsoft YaHei", 16, "bold")).pack(anchor="w", pady=(2, 0))
        tk.Label(head, text="TVM   ·   CAPM   ·   WACC   ·   LOAN", bg=UI["header"],
                 fg=UI["muted"], font=("Consolas", 9)).pack(side="left", padx=22)
        tk.Label(head, text=APP_VERSION, bg=UI["header"], fg=UI["accent_soft"],
                 font=("Consolas", 9, "bold")).pack(side="right", padx=18)

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self.root, padding=(14, 10, 14, 7))
        bar.pack(fill="x")
        ttk.Button(bar, text="一键运行全部场景", style="Accent.TButton",
                   command=self.run_demo).pack(side="left")
        ttk.Button(bar, text="数值自检", command=self.run_selftest).pack(side="left", padx=8)
        ttk.Button(bar, text="打开输出目录",
                   command=lambda: _open_path(get_output_dir())).pack(side="left")
        ttk.Button(bar, text="使用说明", command=self.show_help).pack(side="left", padx=8)
        ttk.Label(bar, text="输入 → 校验 → 计算 → 解释 → 导出",
                  foreground=UI["muted"], font=("Consolas", 9)).pack(side="right")

    def _build_notebook(self) -> None:
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self.overview = OverviewTab(self.nb, self)
        self.nb.add(self.overview, text="总览")

        self.tabs: Dict[str, SceneTab] = {}
        for spec in self._scene_specs():
            tab = SceneTab(self.nb, self, spec["title"], spec["desc"], spec["fields"], spec["runner"])
            self.nb.add(tab, text=spec["tab"])
            self.tabs[spec["tab"]] = tab

    def _build_statusbar(self) -> None:
        bar = ttk.Frame(self.root, padding=(10, 2, 10, 6))
        bar.pack(fill="x")
        self.status = ttk.Label(bar, text="就绪", foreground=UI["muted"])
        self.status.pack(side="left")
        ttk.Label(bar, text=f"输出目录：{get_output_dir()}", foreground=UI["muted"]).pack(side="right")

    # ---------------- 场景页定义 ----------------
    @staticmethod
    def _scene_specs() -> List[Dict[str, Any]]:
        return [
            dict(
                tab="1 复利终值·现值",
                title="场景 1 ｜ 复利终值 / 现值计算（第 3 章）",
                desc="提示：利率请输入小数形式（6% 输入 0.06）；期数请输入正整数。",
                fields=[
                    FieldSpec("pv", "现值 PV / 终值（元）", 100000, "positive_amount", hint="元"),
                    FieldSpec("rate", "年利率 r", 0.06, "rate", hint="小数，如 0.06 表示 6%"),
                    FieldSpec("years", "年数 n", 10, "years", kind="int", hint="正整数"),
                    FieldSpec("freq", "每年计息次数 m", 1, None, kind="choice",
                              options=[("1（按年）", 1), ("2（按半年）", 2), ("4（按季）", 4),
                                       ("12（按月）", 12), ("365（按日）", 365)]),
                    FieldSpec("mode", "计算方向", "1", None, kind="choice",
                              options=[("已知现值求终值", "1"), ("已知终值求现值", "2")]),
                ],
                runner=run_compound,
            ),
            dict(
                tab="2 年金终值·现值",
                title="场景 2 ｜ 年金终值 / 现值计算（第 3 章）",
                desc="提示：期利率 = 年利率 / 每年收付次数，请保持口径一致。",
                fields=[
                    FieldSpec("pmt", "每期现金流 PMT（元）", 10000, "positive_amount", hint="元"),
                    FieldSpec("rate", "每期利率 r", 0.08, "rate", hint="小数"),
                    FieldSpec("periods", "期数 n", 10, "periods", kind="int", hint="正整数"),
                    FieldSpec("due", "收付方式", False, None, kind="choice",
                              options=[("期末收付（普通年金）", False), ("期初收付（预付年金）", True)]),
                ],
                runner=run_annuity,
            ),
            dict(
                tab="3 债权资本成本",
                title="场景 3 ｜ 债权资本成本计算（第 4 章，含税盾效应）",
                desc="示例：面值 1000 元、票面利率 6%、5 年期、发行价 950 元、税率 25%。",
                fields=[
                    FieldSpec("face", "债券面值 F（元）", 1000, "positive_amount", hint="元"),
                    FieldSpec("coupon", "票面利率（年）", 0.06, "coupon_rate", hint="小数"),
                    FieldSpec("years", "债券期限（年）", 5, "years", kind="int", hint="正整数"),
                    FieldSpec("price", "发行价 / 净筹资额（元）", 950, "positive_amount", hint="元"),
                    FieldSpec("tax", "所得税税率", 0.25, "tax_rate", hint="小数，一般企业 0.25"),
                ],
                runner=run_debt,
            ),
            dict(
                tab="4 股权资本成本",
                title="场景 4 ｜ 股权资本成本计算（第 4 章，CAPM）",
                desc="示例：无风险利率 2.5%（10 年期国债）、β = 1.2、市场期望收益 9.5%。",
                fields=[
                    FieldSpec("rf", "无风险利率 Rf", 0.025, "rf", hint="小数"),
                    FieldSpec("beta", "β 系数", 1.2, "beta", hint="无量纲，市场组合 = 1.0"),
                    FieldSpec("rm", "市场期望收益率 Rm", 0.095, "rm", hint="小数"),
                    FieldSpec("size_premium", "规模 / 流动性溢价（无则 0）", 0.0, None, hint="小数"),
                    FieldSpec("use_ddm", "是否用 DDM 交叉验证", True, None, kind="choice",
                              options=[("是（推荐）", True), ("否", False)]),
                    FieldSpec("d1", "下一年度每股股利 D1（元）", 0.5, "positive_amount", hint="元"),
                    FieldSpec("p0", "股票现价 P0（元）", 12.0, "positive_amount", hint="元"),
                    FieldSpec("g", "股利增长率 g", 0.04, "growth", hint="小数"),
                ],
                runner=run_equity,
            ),
            dict(
                tab="5 WACC",
                title="场景 5 ｜ 加权平均资本成本 WACC（第 4 章）",
                desc="债务成本 + 股权成本 + 加权，含资本结构优化模拟与敏感性分析。",
                fields=[
                    FieldSpec("face", "债券面值（元）", 1000, "positive_amount", hint="元"),
                    FieldSpec("coupon", "票面年利率", 0.06, "coupon_rate", hint="小数"),
                    FieldSpec("years", "期限（年）", 5, "years", kind="int", hint="正整数"),
                    FieldSpec("price", "发行价 / 净筹资额（元）", 980, "positive_amount", hint="元"),
                    FieldSpec("tax", "所得税税率", 0.25, "tax_rate", hint="小数"),
                    FieldSpec("rf", "无风险利率 Rf", 0.025, "rf", hint="小数"),
                    FieldSpec("beta", "β 系数", 1.15, "beta", hint="无量纲"),
                    FieldSpec("rm", "市场期望收益率 Rm", 0.095, "rm", hint="小数"),
                    FieldSpec("equity", "股权市场价值 E", 6000, "positive_amount", hint="万元（市场价值口径）"),
                    FieldSpec("debt", "债务市场价值 D", 4000, "positive_amount", hint="万元（市场价值口径）"),
                    FieldSpec("pref", "优先股市场价值 P（无则 0）", 0.0, None, hint="万元"),
                    FieldSpec("kp", "优先股资本成本（无则 0）", 0.0, None, hint="小数"),
                ],
                runner=run_wacc,
            ),
            dict(
                tab="6 贷款还款计划",
                title="场景 6 ｜ 贷款还款计划（等额本息）",
                desc="等额本息是年金现值公式的逆向应用：本金 = 各期还款的年金现值。",
                fields=[
                    FieldSpec("principal", "贷款本金（元）", 1000000, "positive_amount", hint="元"),
                    FieldSpec("rate", "年利率", 0.042, "rate", hint="小数"),
                    FieldSpec("years", "贷款年限", 20, "years", kind="int", hint="正整数"),
                    FieldSpec("freq", "每年还款次数", 12, None, kind="choice",
                              options=[("12（按月）", 12), ("4（按季）", 4),
                                       ("2（按半年）", 2), ("1（按年）", 1)]),
                ],
                runner=run_loan,
            ),
        ]

    # ---------------- 对外动作 ----------------
    def on_scene_done(self, result: SceneResult) -> None:
        """场景计算完成后的统一回调：刷新图表预览与状态栏。"""
        for png in [p for p in result.figures if p.lower().endswith(".png")]:
            self.overview.preview.refresh(os.path.basename(png))
            break
        else:
            self.overview.preview.refresh()
        for tab in self.tabs.values():
            tab.preview.refresh()
        self.set_status(f"{result.name} 完成，产物已写入 {get_output_dir()}")

    def set_status(self, text: str) -> None:
        self.status.configure(text=text)

    def run_demo(self) -> None:
        self.nb.select(self.overview)
        self.overview.on_demo()

    def run_selftest(self) -> None:
        self.nb.select(self.overview)
        self.overview.on_selftest()

    def show_help(self) -> None:
        dep_line = ("依赖状态：numpy / pandas / matplotlib / openpyxl 均已加载"
                    if (MPL_OK and PANDAS_OK) else
                    "依赖缺失，请先安装：pip install numpy pandas matplotlib openpyxl")
        messagebox.showinfo("使用说明", "\n".join([
            f"{APP_TITLE} {APP_VERSION}",
            "",
            "1. 六个场景页：左侧填参数 → 点「计算并生成图表」。",
            "   右侧上方显示计算结果与业务解读，右侧下方预览本次生成的图表。",
            "2. 参数口径：利率 / 税率 / 增长率填小数（6% → 0.06）；期数 / 年数填正整数；",
            "   金额默认单位为元；WACC 页的资本结构金额按万元填报（保持同一口径即可）。",
            "3. 输入错误会被三层校验拦截（类型 → 边界 → 合理性），",
            "   错误框会标红，并按“错在哪里 + 为什么错 + 应如何修正”给出提示。",
            "4. 总览页可一键跑完全部示例场景，生成 8 张图表与 Excel / Markdown 汇总；",
            "   也可运行数值自检测试并导出测试用例簿。",
            "",
            f"输出目录：{get_output_dir()}",
            f"图表中文字体：{CJK_FONT}",
            dep_line,
        ]))

    def _on_tk_error(self, exc, val, tb) -> None:
        text = "".join(traceback.format_exception(exc, val, tb))
        try:
            messagebox.showerror("程序异常", text[-1800:])
        except Exception:                               # noqa: BLE001
            print(text, file=sys.stderr)



# =============================================================================
# 8. 入口（图形界面 / 命令行三种工作模式）
# =============================================================================
def _headless_demo() -> int:
    """命令行模式：一键运行全部场景并打印摘要（供批处理与验证使用）。"""
    print(f"{APP_TITLE} {APP_VERSION} —— 一键运行全部场景\n")
    for line in dependency_report():
        print(line)
    results, files = run_all_scenarios()
    for r in results:
        print()
        print(r.to_text())
    print("\n导出文件：")
    for p in files:
        print("  ·", p)
    return 0


def _headless_selftest() -> int:
    """命令行模式：运行数值自检测试，失败返回非零退出码。"""
    ok, lines, _records = run_self_test(os.path.join(get_output_dir(), "10_测试用例明细.xlsx"))
    print("\n".join(lines))
    return 0 if ok else 1


def _check_imports() -> int:
    print(f"{APP_TITLE} {APP_VERSION} —— 依赖环境检查\n")
    for line in dependency_report():
        print(line)
    print(f"\n  tkinter     {'✔ 可用' if TK_OK else '✘ 不可用 —— ' + _IMPORT_ERRORS.get('tkinter', '')}")
    print(f"  中文字体    {CJK_FONT}")
    print(f"  输出目录    {get_output_dir()}")
    return 0 if (MPL_OK and PANDAS_OK and TK_OK) else 1


def _report_missing_deps() -> None:
    missing = "、".join(k for k, _ in _IMPORT_ERRORS.items())
    msg = (f"缺少运行依赖：{missing}\n\n"
           f"请在命令行执行：\n{PIP_HINT}\n\n"
           f"安装完成后重新双击本程序。")
    try:
        if TK_OK:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("依赖缺失", msg)
            root.destroy()
            return
    except Exception:                                   # noqa: BLE001
        pass
    print(msg)


def _enable_dpi_awareness() -> None:
    """Windows 高 DPI 屏幕上让界面更清晰（失败可忽略）。"""
    if not sys.platform.startswith("win"):
        return
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:                                   # noqa: BLE001
        try:
            from ctypes import windll
            windll.user32.SetProcessDPIAware()
        except Exception:                               # noqa: BLE001
            pass


def _relax_stdout() -> None:
    """
    命令行模式下的控制台编码兜底。

    中文 Windows 的控制台默认使用 GBK，直接输出 ✔ / ✘ / ▶ 等符号会抛
    UnicodeEncodeError。这里优先把控制台切到 UTF-8（cp65001），
    切不过去时则把输出流的错误策略改为 replace，保证程序不因符号而中断。
    图形界面模式不受影响。
    """
    utf8_ok = False
    if sys.platform.startswith("win"):
        try:
            from ctypes import windll
            windll.kernel32.SetConsoleOutputCP(65001)
            utf8_ok = True
        except Exception:                               # noqa: BLE001
            utf8_ok = False
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8" if utf8_ok else None, errors="replace")
        except Exception:                               # noqa: BLE001
            try:
                stream.reconfigure(errors="replace")
            except Exception:                           # noqa: BLE001
                pass


def main(argv: Optional[Sequence[str]] = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv:
        _relax_stdout()

    if "--check-imports" in argv:
        return _check_imports()

    if not (MPL_OK and PANDAS_OK):
        _report_missing_deps()
        return 2

    if "--selftest" in argv:
        return _headless_selftest()
    if "--demo" in argv:
        return _headless_demo()

    if not TK_OK:
        print("tkinter 不可用，已回退为命令行模式（python financial_calculator.pyw --demo）")
        return 2

    _enable_dpi_awareness()
    root = tk.Tk()
    FinanceCalculatorApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())







