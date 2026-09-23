# -*- coding: utf-8 -*-
"""
validators.py —— AI 辅助的智能参数校验与业务场景解释模块
============================================================

对应作业要求"（二）AI 辅助功能"：
  1. 参数校验逻辑：利率范围、期数为正整数、比重合计为 1、β 合理区间等，
     在计算前拦截非法输入，避免"算出结果但结果无意义"。
  2. 业务场景解释：把抽象数字翻译成管理层/业务方听得懂的话，
     例如"当前 WACC 为 8.5%，低于行业平均 10%，说明企业融资成本较低"。

设计思路
--------
校验分层：
  L1 类型/缺失校验  -> 用户是否填了、是不是数字
  L2 硬性边界校验    -> 违反数学定义域，直接报错（如 n <= 0）
  L3 业务合理性校验  -> 数值合法但"反常识"，给出警告（WARN）而非阻断，
                        例如 β = 3 虽可算，但属于极端值需提示复核
校验结果统一返回 ValidationResult，包含 error / warning / hint 三类信息，
使"报错"本身也变成一种教学输出。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

__all__ = [
    "ValidationResult",
    "validate",
    "RULES",
    "check_wacc_weights",
    "check_beta_sanity",
    "interpret_wacc",
    "interpret_tvm",
    "interpret_debt",
    "interpret_equity",
]


# =============================================================================
# 校验结果容器
# =============================================================================
@dataclass
class ValidationResult:
    """一次参数校验的完整结果。"""
    ok: bool = True
    errors: List[str] = field(default_factory=list)      # 阻断级：不修正就无法计算
    warnings: List[str] = field(default_factory=list)    # 提示级：可以算，但需复核
    hints: List[str] = field(default_factory=list)       # 教学级：口径说明

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


# =============================================================================
# 校验规则定义（声明式，便于扩展）
# =============================================================================
# 参数 -> 适用规则。区间单位为"输入值本身的单位"（利率为小数）。
RULES: Dict[str, Dict[str, Any]] = {
    # ---------- 利率类（小数形式，0.05 = 5%） ----------
    "rate": {
        "label": "利率",
        "hard": (-0.99, 1.00),        # 单期最高 100%，超过几乎必为口径错误
        "typical": (-0.05, 0.30),     # 常见区间：-5% ~ 30%
        "unit": "小数（0.05 表示 5%）",
    },
    "discount_rate": {
        "label": "折现率",
        "hard": (0.0, 1.00),
        "typical": (0.01, 0.30),
        "unit": "小数",
    },
    "rf": {
        "label": "无风险利率",
        "hard": (0.0, 0.20),
        "typical": (0.015, 0.06),
        "unit": "小数；中国常用 10 年期国债收益率 2%~3%",
    },
    "rm": {
        "label": "市场期望收益率",
        "hard": (-0.5, 0.60),
        "typical": (0.05, 0.20),
        "unit": "小数；长期股票市场年化收益经验值 8%~12%",
    },
    "beta": {
        "label": "β系数",
        "hard": (0.0, 5.0),
        "typical": (0.3, 2.5),
        "unit": "无量纲；市场组合定义为 1.0",
    },
    "tax_rate": {
        "label": "所得税税率",
        "hard": (0.0, 0.99),
        "typical": (0.0, 0.35),
        "unit": "小数；中国一般企业所得税 25%，高新 15%",
    },
    "coupon_rate": {
        "label": "票面利率",
        "hard": (0.0, 0.50),
        "typical": (0.01, 0.12),
        "unit": "小数",
    },
    "growth": {
        "label": "增长率",
        "hard": (-0.5, 1.0),
        "typical": (-0.1, 0.15),
        "unit": "小数",
    },
    # ---------- 期数类 ----------
    "periods": {
        "label": "期数",
        "hard": (1, 1200),
        "typical": (1, 600),
        "unit": "期（正整数值）",
        "must_be_int": True,
    },
    "years": {
        "label": "年数",
        "hard": (1, 100),
        "typical": (1, 50),
        "unit": "年（正整数值）",
        "must_be_int": True,
    },
    # ---------- 金额类 ----------
    "amount": {
        "label": "金额",
        "hard": (0.0, 1e15),
        "typical": (0.0, 1e12),
        "unit": "元",
    },
    "positive_amount": {
        "label": "金额（须为正）",
        "hard": (1e-9, 1e15),
        "typical": (0.0, 1e12),
        "unit": "元",
    },
    "weight": {
        "label": "权重",
        "hard": (0.0, 1.0),
        "typical": (0.0, 1.0),
        "unit": "小数（0~1 之间）",
    },
}


def validate(field_name: str, value: Any, rules: Dict[str, Any] = None) -> ValidationResult:
    """
    对单个参数按 RULES 中同名规则做三层校验。

    参数
    ----
    field_name : 参数名（需存在于 RULES，否则返回"无规则"提示）
    value      : 待校验值

    返回
    ----
    ValidationResult
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


def check_wacc_weights(wd: float, we: float, wp: float = 0.0,
                       tol: float = 1e-6) -> ValidationResult:
    """
    WACC 权重合计校验：三者之和必须等于 1（或同口径金额，由调用方确认）。
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


# =============================================================================
# 业务场景解释引擎（结果 → 通俗语言）
# =============================================================================
def interpret_wacc(w: Dict[str, float], industry_avg: float = 0.10,
                   roic: Optional[float] = None,
                   rf: float = 0.025) -> List[str]:
    """
    自动生成 WACC 结果的业务解读。

    解读维度：
      1. 绝对水平 —— 与行业平均对比，判断融资成本高低
      2. 相对水平 —— 与无风险利率的利差，即承担的总体风险溢价
      3. 结构成因 —— 股权/债权各自对 WACC 的贡献度
      4. 价值创造 —— 若提供 ROIC，判断企业是否创造经济增加值（EVA > 0）
    """
    out: List[str] = []
    v = w["WACC"]
    gap = v - industry_avg

    if gap <= -0.02:
        t = "明显低于"
        s = "融资成本具有显著优势，同等项目下可承受更低的投资回报率，安全边际较大。"
    elif gap <= -0.005:
        t = "低于"
        s = "融资成本处于较优水平，在项目投资决策中具有相对竞争力。"
    elif gap < 0.005:
        t = "基本持平于"
        s = "融资成本与行业平均水平接近，投资决策应更依赖项目的差异化优势。"
    elif gap < 0.02:
        t = "略高于"
        s = "融资成本略高，需通过提升项目回报率或优化资本结构来弥补。"
    else:
        t = "明显高于"
        s = "融资成本偏高，建议重点审视：高 β 带来的股权成本、财务风险推高的债务成本，以及资本结构是否过度依赖股权。"

    out.append(f"当前 WACC 为 {v:.2%}，{t}行业平均水平（{industry_avg:.2%}，差值 {gap:+.2%}），{s}")

    spread = v - rf
    out.append(
        f"相对无风险利率（{rf:.2%}）的风险补偿为 {spread:.2%}，"
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
            out.append(
                f"企业投入资本回报率 ROIC = {roic:.2%}，高于 WACC {eva:.2%}，"
                f"说明当前经营正在创造经济增加值（EVA > 0），股东价值在增加。"
            )
        else:
            out.append(
                f"企业投入资本回报率 ROIC = {roic:.2%}，低于 WACC {abs(eva):.2%}，"
                f"说明经营回报未能覆盖全部资本成本（EVA < 0），长期看会侵蚀股东价值，"
                f"需通过提价、降本或压缩低效资本占用改善。"
            )
    return out


def interpret_tvm(
    kind: str,
    inputs: Dict[str, float],
    result: float,
    *,
    periods_per_year: Optional[int] = None,
    period_unit: str = "年",
) -> List[str]:
    """
    复利/年金结果的时间价值解读。

    口径约定（二选一，必须由调用方显式声明，函数内部不做隐性换算）
    ------------------------------------------------------------
    * ``periods_per_year is None``（默认）——``inputs["rate"]`` 是**年化利率**，
      ``inputs["periods"] / inputs["years"]`` 视为**年数**。只有该口径下才允许
      出现"72 法则""每 N 年翻一番""N 年内的增长倍数"这类表述。
    * ``periods_per_year = m`` —— ``inputs["rate"]`` 是**期利率**（每期利率），
      期数为 ``inputs["periods"]``。此时解读必须同时给出"期利率 / 期数 / 折合年化
      利率 / 折合年数"四个量，并声明口径；**不得套用 72 法则**——72 法则的成立
      前提是年化复利口径，把它套在期利率上会产生量级错误（例如月利率 0.67%
      会被说成"年化 0.67%"）。

    参数
    ----
    kind             : 场景名，如 "复利终值" / "年金现值" / "预付年金现值"
    inputs           : {"rate": 利率, "periods" 或 "years": 期数/年数, "amount": 本金或每期金额}
    result           : 计算得到的现值/终值，用于计算占比
    periods_per_year : 每年计息/收付次数；传入即表示 rate 为期利率
    period_unit      : 期利率口径下每期的自然语言单位（"月" / "季" / "年"）

    返回
    ----
    逐行解读文本
    """
    out: List[str] = []
    r = inputs.get("rate")
    n = inputs.get("periods") or inputs.get("years")

    if r is not None and n is not None:
        if periods_per_year is None:
            # ---------- 年化口径：允许 72 法则 ----------
            if r > 0.02:
                out.append(
                    f"在年化利率 {r:.2%} 下，资金约每 {72 / (r * 100):.1f} 年翻一番"
                    f"（按“72 法则”快速估算），{n:.0f} 年内的增长倍数约为 {(1 + r) ** n:.2f} 倍。"
                )
            elif r > 0:
                # 72 法则仅对约 5%~20% 的年利率近似良好；低利率区间误差迅速放大
                exact = math.log(2) / math.log1p(r)
                out.append(
                    f"年化利率 {r:.2%} 偏低，“72 法则”在该区间误差较大（估算 "
                    f"{72 / (r * 100):.0f} 年 vs 精确 {exact:.1f} 年），不再套用；"
                    f"{n:.0f} 年内的复利增长倍数约为 {(1 + r) ** n:.2f} 倍。"
                )
            else:
                # r <= 0 时不存在"翻番"概念，72 法则不成立
                out.append(
                    f"年化利率为负（{r:.2%}），资金 {n:.0f} 年后缩水至本金的 {(1 + r) ** n:.2%}；"
                    f"此情形下不存在“翻番”概念，“72 法则”不成立。负利率对应通货紧缩或"
                    f"资产实际贬值的极端情形，解读时应关注购买力而非名义金额。"
                )
        else:
            # ---------- 期利率口径：先声明口径，不做隐性换算 ----------
            m = int(periods_per_year)
            ear = (1.0 + r) ** m - 1.0
            total_years = n / m
            out.append(
                f"口径声明：以下按【期利率口径】解读——每期利率 {r:.4%}（按{period_unit}计息）、"
                f"共 {n:.0f} 期，折合年数 {total_years:.2f} 年，"
                f"折合年化有效利率（EAR）{(1.0 + r) ** m - 1.0:.2%}（而非 {r:.4%}）。"
                f"“72 法则”只适用于年化口径，此处不套用。"
            )
            out.append(
                f"期内的累计增长倍数为 {(1.0 + r) ** n:.4f} 倍"
                f"（= (1 + {r:.4%})^{n:.0f}），这是 {n:.0f} 期复利后的总倍数。"
            )

    if kind.startswith("复利终值"):
        pv = inputs.get("amount", 0)
        if pv:
            out.append(
                f"期初投入 {pv:,.2f} 元，到期金额 {result:,.2f} 元，"
                f"其中利息收益 {result - pv:,.2f} 元，占比 {(result - pv) / pv:.1%}。"
            )
    elif kind.startswith("复利现值"):
        fv = inputs.get("amount", 0)
        if fv:
            out.append(
                f"未来可收回 {fv:,.2f} 元，按上述口径折现后今天的价值为 {result:,.2f} 元，"
                f"折现掉的部分 {fv - result:,.2f} 元（{(fv - result) / fv:.1%}）就是等待的时间成本。"
            )
    elif "年金" in kind:
        pmt = inputs.get("amount", 0)
        total = pmt * (n or 0)
        unit = period_unit if periods_per_year is not None else "年"
        out.append(
            f"累计投入 {total:,.2f} 元，{'现值' if '现值' in kind else '终值'}为 {result:,.2f} 元；"
            f"每期收付发生在{'期初（预付年金，比普通年金多赚一期利息）' if '预付' in kind else '期末（普通年金）'}，"
            f"期数单位为“{unit}”。"
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
    if market_rate is not None:
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
        out.append(
            f"β = {beta:.2f} > 1，说明该股票波动大于市场整体（如周期性制造、券商、房地产），"
            f"投资者要求更高的风险补偿，股权融资相对更“贵”。"
        )
    elif beta < 1:
        out.append(
            f"β = {beta:.2f} < 1，属于防御型资产（如公用事业、食品饮料），"
            f"与市场联动较弱，股权成本相对较低，适合作为高杠杆企业的资本结构调整方向。"
        )
    else:
        out.append("β ≈ 1，风险特征与市场整体同步。")
    out.append(
        "提示：CAPM 的结果对无风险利率与风险溢价假设高度敏感，"
        "建议同时用股利折现模型（DDM）或可比公司法交叉验证。"
    )
    return out
