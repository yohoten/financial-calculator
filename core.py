# -*- coding: utf-8 -*-
"""
core.py —— 货币时间价值与资本成本核心计算引擎
================================================

对应课程知识点
--------------
* 第 3 章 货币时间价值
    - 复利终值 FV = PV * (1 + r)^n
    - 复利现值 PV = FV / (1 + r)^n
    - 普通年金终值 FVA = PMT * [((1+r)^n - 1) / r]
    - 普通年金现值 PVA = PMT * [(1 - (1+r)^-n) / r]
    - 预付年金（先付年金）终值/现值 = 普通年金结果 * (1 + r)
    - 永续年金现值 PVP = PMT / r
* 第 4 章 资本成本
    - 债务资本成本（税前）: 用现金流量法（IRR）精确求解，或近似公式
    - 债务资本成本（税后）: Kd_after = Kd_before * (1 - T)   ← 税盾效应
    - 股权资本成本（CAPM）: Ks = Rf + beta * (Rm - Rf) [+ 规模溢价 SP]
    - 加权平均资本成本 WACC = wd * Kd_after + ws * Ks + wp * Kp

设计原则
--------
1. 纯函数：所有计算函数只依赖入参，不做输入输出、不 print，方便单元测试与复用。
2. 双引擎互验：每个时间价值指标都提供"手写公式实现"与"numpy 财务函数实现"两条路径，
   二者互为校验，规避 Excel 手工公式易出错的问题。
3. 利率一律使用"期间利率"（r）与"期数"（n）配对；年化利率与年数通过
   periods_per_year 换算，避免 r/n 口径错配这一最常见的错误。

作者: WorkBuddy  |  版本: 1.0.0
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Sequence

import numpy as np

__all__ = [
    "FinanceError",
    "fv_compound",
    "pv_compound",
    "fva_ordinary",
    "pva_ordinary",
    "fva_due",
    "pva_due",
    "pv_perpetuity",
    "fv_from_schedule",
    "cost_of_debt",
    "cost_of_debt_approx",
    "cost_of_debt_after_tax",
    "capm_required_return",
    "cost_of_equity_ddm",
    "wacc",
    "bond_price",
    "npv",
    "irr",
    "amortization_schedule",
]


class FinanceError(ValueError):
    """财务计算域异常：参数超出数学模型的有效定义域。"""


# =============================================================================
# 0. 通用工具
# =============================================================================
def _check_periods(n: float) -> float:
    """期数校验：必须为正的有限数。"""
    if not np.isfinite(n) or n <= 0:
        raise FinanceError(f"期数 n 必须为大于 0 的有限数，当前为 {n}。")
    return float(n)


def _check_rate(r: float) -> float:
    """利率校验：必须为大于 -100% 的有限数（-100% 意味着本金归零，无经济意义）。"""
    if not np.isfinite(r) or r <= -1:
        raise FinanceError(f"期间利率 r 必须大于 -100%，当前为 {r}。")
    return float(r)


# =============================================================================
# 1. 复利终值 / 现值（第 3 章）
# =============================================================================
def pv_compound(fv: float, r: float, n: float, m: int = 1) -> float:
    """
    复利现值：PV = FV / (1 + r/m)^(n*m)

    参数
    ----
    fv : 终值（未来某期金额）
    r  : 年名义利率（小数，如 6% 传 0.06）
    n  : 年数
    m  : 每年计息次数（1=年计息，2=半年，4=季，12=月，365=日）

    返回
    ----
    现值 PV

    示例
    ----
    >>> round(pv_compound(1_000_000, 0.06, 5, 12), 2)   # 年化 6%、月复利、5 年后的 100 万现值
    741372.92
    """
    _check_periods(n)
    _check_rate(r / m)
    return float(fv) / (1.0 + r / m) ** (n * m)


def fv_compound(pv: float, r: float, n: float, m: int = 1) -> float:
    """
    复利终值：FV = PV * (1 + r/m)^(n*m)

    参数含义同 pv_compound。m 越大（计息越频繁），实际年利率（EAR）越高，
    这一效应称为"复利频率效应"。

    >>> round(fv_compound(100_000, 0.05, 10, 1), 2)    # 10 万，年利率 5%，年复利，10 年
    162889.46
    >>> round(fv_compound(100_000, 0.05, 10, 12), 2)   # 同样条件，改为月复利
    164700.95
    """
    _check_periods(n)
    _check_rate(r / m)
    return float(pv) * (1.0 + r / m) ** (n * m)


def effective_annual_rate(r: float, m: int = 1) -> float:
    """
    名义年利率 → 实际年利率（EAR / EFF）：
        EAR = (1 + r/m)^m - 1
    用于量化"计息频率"对真实资金成本的影响。
    """
    _check_rate(r / m)
    return (1.0 + r / m) ** m - 1.0


# =============================================================================
# 2. 年金终值 / 现值（第 3 章）
# =============================================================================
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
    预付年金终值：
        FVA_due = FVA_ordinary * (1 + r)
    含义：期初存入比期末存入多赚一期利息，故终值放大 (1+r) 倍。
    """
    return fva_ordinary(pmt, r, n) * (1.0 + _check_rate(r))


def pv_perpetuity(pmt: float, r: float, growth: float = 0.0) -> float:
    """
    永续年金现值（含固定增长模型，戈登模型）：
        PVP = PMT / (r - g)
    r <= g 时公式无经济意义（级数不收敛），抛出异常。
    """
    r = _check_rate(r)
    if r <= growth:
        raise FinanceError(f"折现率 r({r:.4%}) 必须大于增长率 g({growth:.4%})，否则永续年金现值不收敛。")
    return float(pmt) / (r - growth)


# =============================================================================
# 3. 自定义现金流时间轴
# =============================================================================
def fv_from_schedule(cashflows: Sequence[float], r: float) -> Dict[str, object]:
    """
    给定不规则现金流序列，逐期滚动计算终值：
        FV_t = FV_(t-1) * (1 + r) + CF_t

    参数
    ----
    cashflows : 各期现金流序列，第 0 期为期初（可为负=投入，正=收回）
    r         : 每期利率

    返回
    ----
    dict: {"fv": 终值, "path": [各期滚动终值], "total_in": 累计投入, "total_gain": 累计收益}
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


# =============================================================================
# 4. 债权资本成本（第 4 章）
# =============================================================================
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

    参数
    ----
    face_value  : 债券面值（到期偿还额）
    coupon_rate : 票面年利率（小数）
    years       : 债券期限（年）
    price       : 发行价/市价（净筹资额）
    tax_rate    : 所得税税率（用于计算税后成本，体现税盾效应）
    freq        : 每年付息次数（1=年付，2=半年付）

    返回
    ----
    dict: 税前资本成本(期间)、税前年化资本成本、税后年化资本成本、税盾节省额
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
        Kd(税后)  = Kd(税前) * (1 - T)

    分子 = 年票面利息 + 年均资本利得(折价摊销)；
    分母 = 面值与发行价的算术平均（近似平均占用资金）。
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
    税后债务资本成本（税盾效应的一般形式）：
        Kd_after = Kd_before * (1 - T)

    经济含义：债务利息可在企业所得税前扣除，产生"税盾（tax shield）"，
    使债务的实际资本成本低于名义利率。这是债务融资相对于股权融资的
    核心优势，也是 MM 理论（有税）中杠杆提升企业价值的来源。
    """
    if not 0.0 <= tax_rate < 1.0:
        raise FinanceError(f"所得税税率须在 [0, 1) 区间内，当前为 {tax_rate}。")
    return kd_pretax * (1.0 - tax_rate)


# =============================================================================
# 5. 股权资本成本（第 4 章）
# =============================================================================
def capm_required_return(
    rf: float,
    beta: float,
    rm: float,
    size_premium: float = 0.0,
) -> Dict[str, float]:
    """
    资本资产定价模型（CAPM）估算股权资本成本：

        Ks = Rf + β * (Rm - Rf) + SP

    其中：
        Rf            : 无风险利率（通常取 10 年期国债到期收益率）
        β (beta)      : 系统性风险系数，β=1 与市场同步，β>1 波动放大
        (Rm - Rf)     : 市场风险溢价（MRP）
        SP            : 规模/流动性等附加风险溢价（扩展项，可置 0）

    返回
    ----
    dict: 股权资本成本、风险溢价等明细，便于报告引用。
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
    股利折现模型（DDM / 戈登增长模型）估算股权资本成本：
        Ks = D1 / P0 + g
    适用于分红稳定、增长可预期的成熟企业，可与 CAPM 结果交叉验证。
    """
    if p0 <= 0:
        raise FinanceError("股票现价 P0 必须大于 0。")
    return d1 / p0 + g


# =============================================================================
# 6. 加权平均资本成本 WACC（第 4 章）
# =============================================================================
def wacc(
    equity_value: float,
    debt_value: float,
    ks: float,
    kd_after_tax: float,
    pref_value: float = 0.0,
    kp: float = 0.0,
) -> Dict[str, float]:
    """
    加权平均资本成本：

        WACC = (E/V) * Ks + (D/V) * Kd_after_tax + (P/V) * Kp

    其中 V = E + D + P 为总资本（按市场价值口径更佳）。

    注意：进入 WACC 的债务成本必须是**税后**成本，股权成本因股利不可
          税前列支，不享受税盾，直接用 CAPM 结果。
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


# =============================================================================
# 7. 债券定价 / NPV / IRR（辅助工具）
# =============================================================================
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
    """
    净现值：NPV = Σ CF_t / (1 + rate)^t ，CF_0 通常为负的投资额。
    """
    _check_rate(rate)
    cf = np.asarray(cashflows, dtype=float)
    t = np.arange(len(cf))
    return float(np.sum(cf / (1.0 + rate) ** t))


def irr(cashflows: Sequence[float], guess: float = 0.1,
        tol: float = 1e-10, max_iter: int = 200) -> float:
    """
    内部收益率 IRR —— 自实现牛顿迭代法（不用 numpy_financial 外部依赖）。

    求解 f(r) = NPV(r) = 0。
    牛顿法：r_(k+1) = r_k - f(r_k) / f'(r_k)
    其中 f'(r) = Σ -t * CF_t / (1+r)^(t+1)

    稳健性设计：牛顿法不收敛时，退化为 [-0.99, 10] 区间上的二分法，
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
                          years: int, freq: int = 12) -> "Dict[str, object]":
    """
    生成等额本息还款计划表（房贷/车贷场景）。

    每期还款额 PMT 由年金现值公式反解：
        PMT = P * r / [1 - (1 + r)^-n]
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
