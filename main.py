#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main.py —— 货币时间价值与资本成本计算器（命令行交互主程序）
==============================================================

运行方式
--------
    python3 main.py              # 交互式菜单模式
    python3 main.py --demo       # 一键运行全部示例并输出图表与 Excel
    python3 main.py --scene 1    # 直接运行指定场景（1~7）
    python3 main.py --gui        # 启动 tkinter 图形界面

功能菜单
--------
    1. 复利终值 / 现值计算              (第 3 章)
    2. 年金终值 / 现值计算              (第 3 章)
    3. 债权资本成本计算（含税盾效应）    (第 4 章)
    4. 股权资本成本计算（CAPM / DDM）    (第 4 章)
    5. 加权平均资本成本 WACC 计算        (第 4 章)
    6. 贷款还款计划（等额本息）          (第 3 章应用)
    7. 永续年金现值 / 自定义现金流序列    (第 3 章)
    8. 一键运行全部示例场景（生成图表 + Excel）
    0. 退出

设计说明
--------
* 交互层（input/output）与计算层（core.py）完全解耦，便于测试与二次开发。
* 每次计算前都经过 validators 三层校验（错误/警告/口径提示），非法输入会被拦截。
* 每次计算后都会输出"业务解读"，把数字翻译成决策语言（AI 辅助功能要求）。
"""

from __future__ import annotations

import argparse
import os
import sys
import traceback
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd

# 允许脚本在任意目录下直接运行
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import core
import validators as vd
import visualize as vz
import report as rp

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUT_DIR, exist_ok=True)

# 已展示过的口径提示（hints）：每类参数只提示一次，避免每次输入都刷屏
_HINT_SHOWN: set = set()

# -----------------------------------------------------------------------------
# 终端输出美化（ANSI 颜色，Windows 终端若不支持可自动降级）
# -----------------------------------------------------------------------------
def _stdout_is_tty() -> bool:
    """安全探测标准输出是否为终端。

    pythonw / .pyw 启动时 sys.stdout 为 None，直接调用 isatty() 会抛异常，
    因此需要兼容——否则图形界面入口一导入就崩。
    """
    try:
        return bool(sys.stdout.isatty())
    except Exception:                                   # noqa: BLE001
        return False


_USE_COLOR = _stdout_is_tty() and os.environ.get("NO_COLOR") is None


def _c(text: str, code: str) -> str:
    return f"\033[{code}m{text}\033[0m" if _USE_COLOR else text


def title(text: str) -> None:
    """打印一级标题。"""
    line = "═" * 62
    print(f"\n{_c(line, '36')}\n  {_c(text, '1;36')}\n{_c(line, '36')}")


def sub(text: str) -> None:
    print(f"\n{_c('▶ ' + text, '33')}")


def ok(text: str) -> None:
    print(f"{_c('  ✔ ', '32')}{text}")


def warn(text: str) -> None:
    print(f"{_c('  ⚠ ', '33')}{text}")


def err(text: str) -> None:
    print(f"{_c('  ✘ ', '31')}{text}")


def info(text: str) -> None:
    print(f"    {text}")


def kv_table(data: Dict[str, Any], title_: str = "", pct_keys: set = None,
             money_keys: set = None) -> None:
    """
    以对齐表格形式打印键值对结果。
    pct_keys  : 需按百分比格式化的键
    money_keys: 需按金额千分位格式化的键
    """
    pct_keys, money_keys = pct_keys or set(), money_keys or set()
    if title_:
        print(f"\n    {_c(title_, '1;37')}")
    width = max(len(str(k)) for k in data) if data else 10
    print("    " + "─" * (width + 26))
    for k, v in data.items():
        if k in pct_keys and isinstance(v, (int, float)):
            val = f"{v:.3%}" if abs(v) < 10 else f"{v:.3f}"
        elif k in money_keys and isinstance(v, (int, float)):
            val = f"{v:,.2f}"
        elif isinstance(v, float):
            val = f"{v:,.4f}"
        else:
            val = str(v)
        print(f"    {str(k).ljust(width)}  {_c(val.rjust(18), '36')}")
    print("    " + "─" * (width + 26))


def _show_hints_once(rule: str, res: vd.ValidationResult) -> None:
    """三层校验的第三层——口径提示（hint）。

    hints 是"教学级"信息（例如'利率须以小数输入，0.05 表示 5%'），
    与 error/warning 不同：它不阻断、也不需每次重复。策略是同一类参数
    在整个会话中只提示一次，既让第三层校验真正生效，又不至于每轮输入都刷屏。
    """
    if rule in _HINT_SHOWN:
        return
    _HINT_SHOWN.add(rule)
    for h in res.hints:
        info(h)


def ask(prompt: str, default: Optional[float] = None, rule: Optional[str] = None) -> float:
    """
    带校验的输入函数：循环直至通过 validators 三层校验。
    输入 "q" 可放弃当前项并抛出 KeyboardInterrupt 返回主菜单。
    """
    while True:
        hint = f"（默认 {default}）" if default is not None else ""
        raw = input(f"    {prompt}{hint}: ").strip()
        if raw.lower() in ("q", "quit", "exit"):
            raise KeyboardInterrupt
        if raw == "" and default is not None:
            return float(default)
        if rule:
            r = vd.validate(rule, raw)
            if not r.ok:
                for e in r.errors:
                    err(e)
                continue
            for w in r.warnings:
                warn(w)
            _show_hints_once(rule, r)
            return float(raw)
        try:
            return float(raw)
        except ValueError:
            err(f"输入 “{raw}” 不是有效数字，请重新输入（示例：0.06 表示 6%）。")


# =============================================================================
# 场景 1：复利终值 / 现值
# =============================================================================
def scene_compound(interactive: bool = True,
                   preset: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """复利终值/现值计算 + 复利增长曲线 + 结果解读。"""
    title("场景 1 | 复利终值与现值计算（第 3 章）")
    if interactive:
        print("    提示：利率请输入小数形式，例如 6% 输入 0.06；期数请输入正整数。")
        pv = ask("现值 PV（元）", 100000, "positive_amount")
        r = ask("年利率 r", 0.06, "rate")
        n = ask("年数 n", 10, "years")
        m = int(ask("每年计息次数 m（1=年/2=半年/4=季/12=月）", 1, None))
        mode = input("    计算方向（1=已知现值求终值，2=已知终值求现值）[1]: ").strip() or "1"
    else:
        p = preset or {}
        pv, r, n, m = p.get("pv", 100000), p.get("rate", 0.06), p.get("years", 10), p.get("freq", 1)
        mode = p.get("mode", "1")

    if mode == "2":
        res = core.pv_compound(pv, r, n, m)
        main_label = "复利现值 PV"
        kv = {"终值 FV": pv, "复利现值 PV": res}
    else:
        res = core.fv_compound(pv, r, n, m)
        main_label = "复利终值 FV"
        kv = {"现值 PV": pv, "复利终值 FV": res}

    ear = core.effective_annual_rate(r, m)
    kv.update({
        "年名义利率": r, "期数(年)": n, "每年计息次数": m,
        "实际年利率 EAR": ear, "利息收益": res - pv if mode == "1" else pv - res,
    })
    kv_table(kv, f"计算结果 · {main_label}",
             pct_keys={"年名义利率", "实际年利率 EAR"},
             money_keys={"现值 PV", "复利终值 FV", "终值 FV", "复利现值 PV", "利息收益"})

    # ---- 业务解读 ----
    sub("业务解读")
    for line in vd.interpret_tvm("复利终值" if mode == "1" else "复利现值",
                                 {"rate": r, "years": n, "amount": pv}, res):
        info(line)
    if m > 1:
        base = core.fv_compound(pv, r, n, 1)
        info(f"计息频率效应：年复利终值为 {base:,.2f} 元，{m} 次复利后提升至 {res:,.2f} 元，"
             f"多出 {res - base:,.2f} 元（{(res / base - 1):.3%}），这就是 EAR 高于名义利率的原因。")

    # ---- 可视化 ----
    fig_path = os.path.join(OUT_DIR, "01_复利增长曲线.png")
    vz.plot_compound_growth(pv, r, int(n), fig_path, freq=int(m))
    ok(f"复利增长曲线已生成：{os.path.relpath(fig_path, BASE_DIR)}")

    return {"场景": "复利终值/现值", "现值PV": pv, "年利率": r, "年数": n,
            "计息次数": m, "实际年利率EAR": ear, main_label: res}


# =============================================================================
# 场景 2：年金终值 / 现值
# =============================================================================
# 期数单位映射：把"每年收付次数"翻译成业务语言，供解读引擎声明口径使用
_PERIOD_UNIT: Dict[int, str] = {1: "年", 2: "半年", 4: "季", 12: "月", 365: "日"}


def scene_annuity(interactive: bool = True,
                  preset: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """普通年金 / 预付年金的终值与现值 + 现金流时间轴 + 折现贡献分解。"""
    title("场景 2 | 年金终值与现值计算（第 3 章）")
    if interactive:
        print("    提示：期利率 = 年利率 / 每年收付次数。本场景的 r 是【期利率】，")
        print("          请一并填写每年收付次数，程序会据此刻画正确的口径（不做隐性换算）。")
        pmt = ask("每期现金流 PMT（元）", 10000, "positive_amount")
        r = ask("每期利率 r", 0.08, "rate")
        n = ask("期数 n", 10, "periods")
        ppy = int(ask("每年收付次数（1=年/2=半年/4=季/12=月）", 1, None))
        due = (input("    收付方式（1=期末/普通年金，2=期初/预付年金）[1]: ").strip() or "1") == "2"
    else:
        p = preset or {}
        pmt, r, n = p.get("pmt", 10000), p.get("rate", 0.08), p.get("periods", 10)
        ppy = int(p.get("periods_per_year", 1))
        due = p.get("due", False)

    pva = core.pva_due(pmt, r, n) if due else core.pva_ordinary(pmt, r, n)
    fva = core.fva_due(pmt, r, n) if due else core.fva_ordinary(pmt, r, n)
    total = pmt * n
    ear = (1.0 + r) ** ppy - 1.0

    kv = {"每期现金流 PMT": pmt, "每期利率(口径基准)": r, "期数": n,
          "每年收付次数": ppy, "折合年化有效利率 EAR": ear, "折合年数": n / ppy,
          "收付方式": "预付年金（期初）" if due else "普通年金（期末）",
          "年金现值 PVA": pva, "年金终值 FVA": fva,
          "累计现金流总额": total, "现值/累计比": pva / total,
          "终值/累计比": fva / total}
    kv_table(kv, "计算结果",
             pct_keys={"每期利率(口径基准)", "折合年化有效利率 EAR", "现值/累计比", "终值/累计比"},
             money_keys={"每期现金流 PMT", "年金现值 PVA", "年金终值 FVA", "累计现金流总额"})

    sub("业务解读")
    for line in vd.interpret_tvm("年金现值" if not due else "预付年金现值",
                                 {"rate": r, "periods": n, "amount": pmt}, pva,
                                 periods_per_year=ppy,
                                 period_unit=_PERIOD_UNIT.get(ppy, "期")):
        info(line)
    if due:
        base_pva = core.pva_ordinary(pmt, r, n)
        base_fva = core.fva_ordinary(pmt, r, n)
        info(f"预付 vs 普通年金：现值相差 {pva - base_pva:,.2f} 元（+{(pva / base_pva - 1):.2%}），"
             f"终值相差 {fva - base_fva:,.2f} 元（+{(fva / base_fva - 1):.2%}）。"
             f"因为每笔款项都提前一期到账/存入，多获得一期利息。")
    info(f"现值仅为累计现金流总额的 {pva / total:.1%}，说明资金的时间价值让远期现金流"
         f"在今天的价值大幅缩水——这正是长期分期付款「总额看起来很多、现值其实不高」的原因。")

    fig_path = os.path.join(OUT_DIR, "02_年金现金流时间轴.png")
    vz.plot_annuity_timeline(pmt, r, int(n), fig_path, due=due,
                             present_value=pva, future_value=fva)
    ok(f"年金现金流时间轴已生成：{os.path.relpath(fig_path, BASE_DIR)}")

    return {"场景": "年金终值/现值", "每期PMT": pmt, "期利率": r, "期数": n, "每年收付次数": ppy,
            "收付方式": "预付" if due else "普通", "年金现值PVA": pva, "年金终值FVA": fva}


# =============================================================================
# 场景 3：债权资本成本
# =============================================================================
def scene_debt(interactive: bool = True,
               preset: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """债权资本成本：IRR 精确法 vs 教材近似法 双引擎互验 + 税盾效应可视化。"""
    title("场景 3 | 债权资本成本计算（第 4 章，含税盾效应）")
    if interactive:
        print("    示例：面值 1000 元、票面利率 6%、5 年期、发行价 950 元、税率 25%")
        fv_ = ask("债券面值 F（元）", 1000, "positive_amount")
        cpn = ask("票面利率（年）", 0.06, "coupon_rate")
        yrs = ask("债券期限（年）", 5, "years")
        price = ask("发行价（元）", 950, "positive_amount")
        tax = ask("所得税税率", 0.25, "tax_rate")
        pay_freq = int(ask("每年付息次数（1=年付/2=半年付/4=季付）", 1, None))
        flot = ask("筹资费率 f（承销费等发行费用占比，无则 0）", 0.0, None)
    else:
        p = preset or {}
        fv_, cpn = p.get("face", 1000), p.get("coupon", 0.06)
        yrs, price, tax = p.get("years", 5), p.get("price", 950), p.get("tax", 0.25)
        pay_freq = int(p.get("pay_freq", 1))
        flot = p.get("flotation", 0.0)

    d = core.cost_of_debt(fv_, cpn, int(yrs), price, tax, pay_freq, flot)
    d_approx = core.cost_of_debt_approx(fv_, cpn, int(yrs), price, tax, flot)

    kv = {"债券面值": fv_, "票面利率": cpn, "期限(年)": yrs, "发行价": price,
          "每年付息次数": pay_freq, "筹资费率 f": flot, "净筹资额": d["净筹资额"],
          "所得税税率": tax, "年利息支出": d["年利息支出"],
          "税前资本成本(YTM法)": d["税前年化资本成本"],
          "税后资本成本": d["税后年化资本成本"],
          "税后成本(严格口径)": d["税后年化资本成本(严格口径)"],
          "税后成本(教材简化)": d["税后年化资本成本(教材简化)"],
          "税盾节省(百分点)": d["税盾节省(百分点)"] / 100,
          "税前资本成本(近似公式法)": d_approx["税前年化资本成本(近似)"],
          "两法差异": d["税前年化资本成本"] - d_approx["税前年化资本成本(近似)"]}
    kv_table(kv, "计算结果",
             pct_keys={"票面利率", "筹资费率 f", "所得税税率", "税前资本成本(YTM法)", "税后资本成本",
                       "税后成本(严格口径)", "税后成本(教材简化)",
                       "税盾节省(百分点)", "税前资本成本(近似公式法)", "两法差异"},
             money_keys={"债券面值", "发行价", "净筹资额", "年利息支出"})

    sub("方法学说明与业务解读")
    info("方法一（主算法）IRR/YTM 法：令债券未来本息现金流现值等于发行价的折现率，"
         "即债权人要求的到期收益率，也是企业真实承担的税前债务成本。")
    info("方法二（校验法）教材近似公式：分母取面值与发行价均值。两法结果高度接近，"
         f"差异仅 {abs(d['税前年化资本成本'] - d_approx['税前年化资本成本(近似)']):.4%}，互为验证。")
    if pay_freq > 1:
        info(f"口径提示：本次为每{_PERIOD_UNIT.get(pay_freq, '期')}付息一次（每年 {pay_freq} 次），"
             f"IRR 求得的是【期间利率】{d['税前期间资本成本']:.4%}，"
             f"程序按有效年利率换算为税前年化 {d['税前年化资本成本']:.4%}"
             f"（而非简单 ×{pay_freq} 得 {d['税前期间资本成本'] * pay_freq:.4%}，后者忽略了年内复利）。")
        info(f"税后成本的三种口径：主口径（EAR×(1-T)）{d['税后年化资本成本']:.4%}、"
             f"严格口径（先对每期利息计税再年化）{d['税后年化资本成本(严格口径)']:.4%}、"
             f"教材简化口径（名义年利率×(1-T)）{d['税后年化资本成本(教材简化)']:.4%}。"
             f"freq = 1 时三者完全相等；freq > 1 时建议在报告中声明所采用的口径。")
    if flot > 0:
        info(f"筹资费用影响：发行价 {price:,.2f} 元扣除 {flot:.2%} 的筹资费用后，"
             f"企业实际可用资金仅 {d['净筹资额']:,.2f} 元，税前债务成本由 "
             f"{core.cost_of_debt(fv_, cpn, int(yrs), price, tax, pay_freq)['税前年化资本成本']:.4%} "
             f"上升到 {d['税前年化资本成本']:.4%}——这就是国内教材在分母中写 “(1 - f)” 的原因。")
    else:
        info("口径边界：本次筹资费率 f = 0，即未计入承销费等发行费用。"
             "国内教材的债务成本公式分母为“发行价 × (1 - f)”；如需与教材例题对齐，"
             "请将筹资费率设为题目给定值（一般为 0 或 2%~5%）。")
    for line in vd.interpret_debt(d, tax, market_rate=cpn):
        info(line)
    if price < fv_:
        info(f"本次债券折价发行（发行价 {price:,.0f} < 面值 {fv_:,.0f}），"
             f"说明市场要求的收益率高于票面利率，投资者以低价买入来补偿利息不足，"
             f"企业的实际债务成本因此高于票面利率。")
    elif price > fv_:
        info("本次债券溢价发行，市场要求的收益率低于票面利率，企业实际债务成本低于票面利率。")
    else:
        info("本次债券平价发行，市场要求收益率等于票面利率，税前债务成本 = 票面利率。")

    fig_path = os.path.join(OUT_DIR, "03_债权资本成本曲线.png")
    vz.plot_debt_cost_curve(fig_path, fv_, cpn, int(yrs), tax_rate=tax,
                            freq=pay_freq, flotation_cost=flot)
    ok(f"债权资本成本曲线已生成：{os.path.relpath(fig_path, BASE_DIR)}")

    return {"场景": "债权资本成本", "面值": fv_, "票面利率": cpn, "期限(年)": yrs,
            "发行价": price, "每年付息次数": pay_freq, "筹资费率": flot,
            "净筹资额": d["净筹资额"], "税率": tax, "税前成本": d["税前年化资本成本"],
            "税后成本": d["税后年化资本成本"],
            "税后成本(严格口径)": d["税后年化资本成本(严格口径)"],
            "税后成本(教材简化)": d["税后年化资本成本(教材简化)"]}


# =============================================================================
# 场景 4：股权资本成本
# =============================================================================
def scene_equity(interactive: bool = True,
                 preset: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """股权资本成本：CAPM 为主，DDM 交叉验证。"""
    title("场景 4 | 股权资本成本计算（第 4 章，CAPM 模型）")
    if interactive:
        print("    示例：无风险利率 2.5%（10 年期国债）、β=1.2、市场期望收益 9.5%")
        rf = ask("无风险利率 Rf", 0.025, "rf")
        beta = ask("β系数", 1.2, "beta")
        rm = ask("市场期望收益率 Rm", 0.095, "rm")
        sp = ask("规模/流动性附加溢价（无则 0）", 0.0, None)
        use_ddm = (input("    是否用股利折现模型（DDM）交叉验证？(1=是 0=否)[1]: ").strip() or "1") == "1"
    else:
        p = preset or {}
        rf, beta, rm = p.get("rf", 0.025), p.get("beta", 1.2), p.get("rm", 0.095)
        sp = p.get("size_premium", 0.0)
        use_ddm = p.get("use_ddm", True)

    e = core.capm_required_return(rf, beta, rm, sp)
    kv = {"无风险利率 Rf": rf, "β系数": beta, "市场期望收益率 Rm": rm,
          "市场风险溢价 MRP": e["市场风险溢价"], "规模溢价": sp,
          "风险补偿部分": e["风险补偿部分"], "股权资本成本 Ks": e["股权资本成本"]}
    kv_table(kv, "CAPM 计算结果",
             pct_keys={"无风险利率 Rf", "市场期望收益率 Rm", "市场风险溢价 MRP",
                       "规模溢价", "风险补偿部分", "股权资本成本 Ks"}, money_keys=set())

    sub("业务解读")
    # β 合理性专项校验（L3 业务合理性层的独立入口，见开发报告 §5.1）
    for msg in vd.check_beta_sanity(beta).warnings:
        warn(msg)
    for line in vd.interpret_equity(e):
        info(line)

    kd_result = None
    if use_ddm:
        info("以下为 DDM（戈登增长模型）交叉验证：")
        if interactive:
            d1 = ask("下一年度每股股利 D1（元）", 0.5, "positive_amount")
            p0 = ask("股票现价 P0（元）", 12.0, "positive_amount")
            g = ask("股利增长率 g", 0.04, "growth")
        else:
            d1 = (preset or {}).get("d1", 0.5)
            p0 = (preset or {}).get("p0", 12.0)
            g = (preset or {}).get("g", 0.04)
        kd_result = core.cost_of_equity_ddm(d1, p0, g)
        kv_table({"每股股利 D1": d1, "股票现价 P0": p0, "增长率 g": g, "DDM 股权成本 Ks": kd_result},
                 "DDM 交叉验证", pct_keys={"增长率 g", "DDM 股权成本 Ks"}, money_keys={"每股股利 D1", "股票现价 P0"})
        diff = kd_result - e["股权资本成本"]
        info(f"两模型差异 {diff:+.3%}：{'差异较小，结论稳健' if abs(diff) < 0.02 else '差异较大，建议进一步核查增长率假设与企业分红政策的可持续性'}。"
             f"实务中通常以 CAPM 结果为主，DDM 用于校验。")

    return {"场景": "股权资本成本", "无风险利率": rf, "β": beta, "市场收益率": rm,
            "股权成本(CAPM)": e["股权资本成本"],
            "股权成本(DDM)": kd_result if kd_result else ""}


# =============================================================================
# 场景 5：WACC
# =============================================================================
def scene_wacc(interactive: bool = True,
               preset: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """完整 WACC 计算：债务成本 + 股权成本 + 加权，含敏感性分析与结构分解图。"""
    title("场景 5 | 加权平均资本成本 WACC 计算（第 4 章）")
    if interactive:
        print("    步骤一：债务资本成本参数")
        fv_ = ask("债券/借款面值（元）", 1000, "positive_amount")
        cpn = ask("票面年利率", 0.06, "coupon_rate")
        yrs = ask("期限（年）", 5, "years")
        price = ask("发行价（元）", 980, "positive_amount")
        tax = ask("所得税税率", 0.25, "tax_rate")
        pay_freq = int(ask("每年付息次数（1=年付/2=半年付）", 1, None))
        flot = ask("筹资费率 f（无则 0）", 0.0, None)
        print("\n    步骤二：股权资本成本参数（CAPM）")
        rf = ask("无风险利率 Rf", 0.025, "rf")
        beta = ask("β系数", 1.15, "beta")
        rm = ask("市场期望收益率 Rm", 0.095, "rm")
        print("\n    步骤三：资本结构（请以市场价值口径填写金额，单位：万元）")
        e_val = ask("股权市场价值 E", 6000, "positive_amount")
        d_val = ask("债务市场价值 D", 4000, "positive_amount")
        p_val = ask("优先股市场价值 P（无则 0）", 0.0, None)
        kp = ask("优先股资本成本（无则 0）", 0.0, None) if p_val > 0 else 0.0
    else:
        p = preset or {}
        fv_, cpn, yrs = p.get("face", 1000), p.get("coupon", 0.06), p.get("years", 5)
        price, tax = p.get("price", 980), p.get("tax", 0.25)
        pay_freq = int(p.get("pay_freq", 1))
        flot = p.get("flotation", 0.0)
        rf, beta, rm = p.get("rf", 0.025), p.get("beta", 1.15), p.get("rm", 0.095)
        e_val, d_val, p_val = p.get("equity", 6000), p.get("debt", 4000), p.get("pref", 0.0)
        kp = p.get("kp", 0.0)

    # ---- 债务成本 ----
    d = core.cost_of_debt(fv_, cpn, int(yrs), price, tax, pay_freq, flot)
    kd_after = d["税后年化资本成本"]
    # ---- 股权成本 ----
    e = core.capm_required_return(rf, beta, rm)
    ks = e["股权资本成本"]
    # ---- WACC：先由 core 层完成计算（core 内部对"资本总额 <= 0"等情形给出中文异常）----
    #     再拿"实际用于计算的权重"去做口径校验，从而保证
    #     "校验用的权重"与"计算用的权重"同源，不会出现两套口径。
    w = core.wacc(e_val, d_val, ks, kd_after, p_val, kp)
    w["_税前债权成本"] = d["税前年化资本成本"]
    w_res = vd.check_wacc_weights(w["债权比重"], w["股权比重"], w["优先股比重"])
    for wmsg in w_res.warnings:
        warn(wmsg)
    for wmsg in w_res.hints:
        info(wmsg)

    kv_table({
        "税前债务成本 Kd(税前)": d["税前年化资本成本"],
        "税后债务成本 Kd(税后)": kd_after,
        "股权成本 Ks(CAPM)": ks,
        "优先股成本 Kp": kp,
        "股权市场价值 E": e_val, "债务市场价值 D": d_val, "优先股 P": p_val,
        "股权权重 wE": w["股权比重"], "债务权重 wD": w["债权比重"],
        "优先股权重 wP": w["优先股比重"],
        "WACC": w["WACC"],
    }, "WACC 汇总结果",
        pct_keys={"税前债务成本 Kd(税前)", "税后债务成本 Kd(税后)", "股权成本 Ks(CAPM)",
                  "优先股成本 Kp", "股权权重 wE", "债务权重 wD", "优先股权重 wP", "WACC"},
        money_keys={"股权市场价值 E", "债务市场价值 D", "优先股 P"})

    sub("业务解读")
    for line in vd.interpret_wacc(w, industry_avg=0.10, roic=0.118, rf=rf):
        info(line)

    sub("资本结构优化：不同负债水平下的 WACC 模拟")
    tax_shield_note = "假设负债上升会小幅推高财务风险，从而抬升 Kd 与 Ks"
    info(tax_shield_note + "；作为对照进行情景测算。")
    info(f"模拟的资本总额保持 WACC 汇总口径的 {w['资本总额']:,.2f} 万元不变；"
         f"基准负债率 wD0 = D/V = {d_val / w['资本总额']:.4%}"
         + (f"，优先股占比 wP = {w['优先股比重']:.4%}（固定不变）" if p_val else "")
         + "。")
    scenarios = simulate_capital_structure(e_val, d_val, ks, d["税前年化资本成本"], tax,
                                           pref=p_val, kp=kp)
    df_s = pd.DataFrame(scenarios)
    print(df_s.to_string(index=False, float_format=lambda x: f"{x:.4f}"))

    # ---- 可视化 ----
    fig1 = os.path.join(OUT_DIR, "04_WACC结构分解.png")
    vz.plot_wacc_breakdown(w, fig1, industry_avg=0.10)
    ok(f"WACC 资本结构分解图：{os.path.relpath(fig1, BASE_DIR)}")

    fig2 = os.path.join(OUT_DIR, "05_WACC敏感性分析.png")
    vz.plot_wacc_sensitivity(fig2, w["股权比重"], w["债权比重"], ks, kd_after)
    ok(f"WACC 敏感性热力图：{os.path.relpath(fig2, BASE_DIR)}")

    fig3 = os.path.join(OUT_DIR, "06_资本结构优化对比.png")
    vz.plot_wacc_scenarios([{"名称": s["负债比例"] + " 负债", "WACC": s["WACC"]} for s in scenarios], fig3)
    ok(f"资本结构优化对比图：{os.path.relpath(fig3, BASE_DIR)}")

    return {"场景": "WACC 计算", "税前债务成本": d["税前年化资本成本"], "税后债务成本": kd_after,
            "股权成本Ks": ks, "股权权重": w["股权比重"], "债务权重": w["债权比重"],
            "优先股权重": w["优先股比重"], "优先股成本": kp, "WACC": w["WACC"],
            "_wacc_dict": w, "_scenarios": scenarios}


def simulate_capital_structure(
    equity: float, debt: float, ks_base: float, kd_pre_base: float, tax: float,
    pref: float = 0.0, kp: float = 0.0,
) -> List[Dict[str, Any]]:
    """
    模拟不同负债比例下的 WACC（资本结构优化）。

    风险传导假设（体现 MM 理论的有税修正 + 财务困境成本）：
        1. 债权成本随负债率上升：Kd = Kd0 * (1 + 0.20 * (wd - wd0))
        2. 股权成本随负债率上升且更敏感：Ks = Ks0 * (1 + 1.05 * (wd - wd0))
           股权剩余索取权在杠杆提高后风险放大更快，故系数显著高于债权。
        3. 同时，负债率超过 55% 后引入"财务困境成本"非线性加速项：
           Ks 额外 + 0.9 * (wd - 0.55)^2 ，Kd 额外 + 0.5 * (wd - 0.55)^2
           模拟评级下调、再融资困难等跳升效应。

    三者共同作用形成 U 型 WACC 曲线：初期税盾效应占优使 WACC 下降，
    越过最优点后财务风险上升的影响超过税盾收益，WACC 转为上升。
    这为资本结构决策提供定量参考——并非负债越多越好。

    口径一致性
    ----------
    资本总额 V = E + D + P 保持不变，优先股金额（因而权重 wp = P/V）固定不变，
    股权权重由 we = 1 - wd - wp 内生决定，WACC 按三来源加权：

        WACC = we * Ks + wd * Kd_税后 + wp * Kp

    因此模拟表的基准负债率 wd0 = D / V 与 `core.wacc()` 给出的债权权重严格同源。
    （修正前此处为 total = equity + debt，遗漏优先股：有优先股时基准 wd0 会
    大于汇总表中的债权权重，导致同一份输出内部口径不一致。）
    """
    total = equity + debt + pref
    if total <= 0:
        raise core.FinanceError("资本总额（股权 + 债权 + 优先股）必须大于 0。")
    wd0 = debt / total
    wp = pref / total

    rows = []
    for wd in np.linspace(0.10, 0.80, 8):
        we = 1.0 - wd - wp
        if we <= 0:                      # 债权 + 优先股合计已超 100%，该情形不成立
            continue
        delta = wd - wd0
        kd_pre = kd_pre_base * (1 + 0.20 * delta)
        ks = ks_base * (1 + 1.05 * delta)

        # 高杠杆区的财务困境成本加速项
        if wd > 0.55:
            stress = (wd - 0.55) ** 2
            ks += 0.9 * stress
            kd_pre += 0.5 * stress

        kd_post = kd_pre * (1 - tax)
        w = we * ks + wd * kd_post + wp * kp
        rows.append({
            "负债比例": f"{wd:.0%}",
            "股权比例": we,
            "税前Kd": kd_pre,
            "税后Kd": kd_post,
            "Ks": ks,
            "WACC": w,
        })
    return rows


# =============================================================================
# 场景 6：贷款还款计划
# =============================================================================
def scene_loan(interactive: bool = True,
               preset: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """等额本息还款计划：把年金现值公式反过来用（第 3 章应用场景）。"""
    title("场景 6 | 贷款还款计划（等额本息，年金公式应用）")
    if interactive:
        principal = ask("贷款本金（元）", 1_000_000, "positive_amount")
        rate = ask("年利率", 0.042, "rate")
        years = ask("贷款年限", 20, "years")
        freq = int(ask("每年还款次数（12=按月）", 12, None))
    else:
        p = preset or {}
        principal, rate, years, freq = (p.get("principal", 1_000_000), p.get("rate", 0.042),
                                        p.get("years", 20), p.get("freq", 12))

    s = core.amortization_schedule(principal, rate, int(years), int(freq))
    pva_check = core.pva_ordinary(s["每期还款额"], rate / freq, years * freq)

    kv_table({"贷款本金": principal, "年利率": rate, "年限": years, "每年还款次数": freq,
              "每期还款额": s["每期还款额"], "总还款额": s["总还款额"],
              "总利息": s["总利息"], "利息/本金比": s["总利息"] / principal,
              "校验：还款额的年金现值": pva_check},
             "还款计划摘要", pct_keys={"年利率", "利息/本金比"},
             money_keys={"贷款本金", "每期还款额", "总还款额", "总利息", "校验：还款额的年金现值"})

    sub("业务解读")
    # 文案随输入参数变化，不再硬编码"每月""20 年""月利率"
    per = _PERIOD_UNIT.get(int(freq), "期")
    n_periods = int(years) * int(freq)
    info(f"每{per}需还款 {s['每期还款额']:,.2f} 元，{int(years)} 年（共 {n_periods} 期）累计还款 "
         f"{s['总还款额']:,.2f} 元，其中利息 {s['总利息']:,.2f} 元，"
         f"相当于本金的 {s['总利息'] / principal:.1%}。")
    info(f"校验逻辑：把每{per}还款额按 {rate / freq:.4%} 的{per}利率折现 {n_periods} 期，"
         f"得到 {pva_check:,.2f} 元，等于贷款本金，验证了等额本息公式的正确性"
         f"（差异 {abs(pva_check - principal):.6f} 元，为浮点误差）。")
    info("这说明等额本息并非「利息少」，而是前期利息占比高、后期本金占比高；"
         "若提前还款，节奏越早，节省的利息越多。")

    # 计划表导出
    df = pd.DataFrame(s["计划表"])
    xlsx = os.path.join(OUT_DIR, "07_贷款还款计划.xlsx")
    with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="还款计划", index=False)
    ok(f"还款计划表已导出：{os.path.relpath(xlsx, BASE_DIR)}（共 {len(df)} 期）")

    # 可视化：本金/利息构成
    fig = os.path.join(OUT_DIR, "08_还款结构.png")
    _plot_loan_structure(df, fig)
    ok(f"还款结构图：{os.path.relpath(fig, BASE_DIR)}")

    return {"场景": "贷款还款计划", "贷款本金": principal, "年利率": rate, "年限": years,
            "月供": s["每期还款额"], "总利息": s["总利息"],
            # 下划线开头为内部字段：不进汇总表，由 report.py 自动导出为「还款计划前24期」工作表
            "_schedule_df": df.head(24)}


def _plot_loan_structure(df: pd.DataFrame, out_path: str) -> str:
    """绘制还款结构图：每期本金/利息堆积 + 剩余本金曲线。"""
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter

    fig, ax = plt.subplots(figsize=(12, 5.2))
    x = df["期数"]
    ax.bar(x, df["利息"], color=vz.C["accent"], width=0.9, label="利息部分")
    ax.bar(x, df["本金"], bottom=df["利息"], color=vz.C["primary"], width=0.9, label="本金部分")
    ax.set_xlabel("期数")
    ax.set_ylabel("每期还款（元）")
    ax.yaxis.set_major_formatter(FuncFormatter(vz._money))
    ax.grid(axis="y", linestyle="--", alpha=0.7)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)

    ax2 = ax.twinx()
    ax2.plot(x, df["剩余本金"], color=vz.C["red"], linewidth=2.2, label="剩余本金")
    ax2.set_ylabel("剩余本金（元）", color=vz.C["red"])
    ax2.yaxis.set_major_formatter(FuncFormatter(vz._money))
    ax2.tick_params(axis="y", colors=vz.C["red"])
    ax2.spines["top"].set_visible(False)

    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="center right", fontsize=9.5)
    ax.set_title("等额本息还款结构：前期利息为主，后期本金为主（剩余本金呈凸曲线下降）", loc="left")
    fig.savefig(out_path)
    plt.close(fig)
    return out_path


# =============================================================================
# 场景 7：永续年金现值 / 自定义现金流序列
# =============================================================================
_DEFAULT_CASHFLOWS: List[float] = [-100_000.0, 30_000.0, 35_000.0, 40_000.0, 45_000.0]


def _parse_cashflows(raw: str) -> List[float]:
    """把 "‑100000,30000,35000" 解析为浮点列表。

    非法输入统一抛出 core.FinanceError，由调用方转成中文提示——
    这样"输入错误"与"计算域错误"走同一条错误通道，不会漏出 Python 堆栈。
    """
    parts = [x.strip() for x in raw.replace("，", ",").split(",") if x.strip()]
    if len(parts) < 2:
        raise core.FinanceError(
            "现金流序列至少需要 2 期（含第 0 期），请用逗号分隔，例如 -100000,30000,35000。")
    try:
        return [float(x) for x in parts]
    except ValueError as ex:
        raise core.FinanceError(
            f"现金流序列中存在无法识别的数值（{ex}）。请只输入数字与逗号，"
            f"投入用负数、收回用正数。") from None


def scene_cashflow(interactive: bool = True,
                   preset: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    永续年金现值（戈登增长模型）与自定义不规则现金流序列。

    直接落地《实践要求（一）》"支持自定义利率、期数、**现金流参数**"：

    * 永续年金：PVP = PMT / (r - g)。r ≤ g 时级数不收敛，
      程序给出"为什么会这样 + 应该怎么改"的中文提示，而不是抛出 Python 异常；
    * 自定义现金流：期数与每期金额均不受限制，按期滚动计算终值
      FV_t = FV_(t-1) × (1 + r) + CF_t，并同时给出折现净现值 NPV 交叉验证。
    """
    title("场景 7 | 永续年金现值与自定义现金流序列（第 3 章）")
    p = preset or {}

    # ---------- A. 永续年金现值 ----------
    sub("A. 永续年金现值 PVP = PMT / (r - g)")
    if interactive:
        print("    提示：g = 0 即标准永续年金；g > 0 即固定增长永续年金（戈登模型）。")
        pmt = ask("每期固定收付金额 PMT（元）", 50000, "positive_amount")
        r = ask("折现率 r（期口径）", 0.08, "rate")
        g = ask("每期增长率 g（无增长填 0）", 0.02, "growth")
    else:
        pmt, r, g = p.get("pmt", 50000), p.get("rate", 0.08), p.get("growth", 0.02)

    pvp: Optional[float] = None
    try:
        pvp = core.pv_perpetuity(pmt, r, g)
    except core.FinanceError as ex:
        warn(f"永续年金现值无法计算：{ex}")

    if pvp is not None:
        kv_table({"每期收付 PMT": pmt, "折现率 r": r, "增长率 g": g,
                  "永续年金现值 PVP": pvp, "资本化倍数 1/(r-g)": 1.0 / (r - g)},
                 "永续年金现值",
                 pct_keys={"折现率 r", "增长率 g"},
                 money_keys={"每期收付 PMT", "永续年金现值 PVP"})
        info(f"含义：只要每期能稳定收到 {pmt:,.0f} 元，在要求回报率 {r:.2%}（且按 {g:.2%} 增长）下，"
             f"这笔永续现金流今天的价值是 {pvp:,.2f} 元，相当于每期金额的 {1 / (r - g):.1f} 倍。")
        info("分母 (r - g) 越小、现值越大，这是低利率环境下高股息资产估值被系统性推高的数学原因；"
             "同时也解释了为什么该模型对 g 极为敏感——g 每提高 1 个百分点，估值成倍放大。")
    else:
        info("提示：永续年金的级数只有在折现率大于增长率时收敛。请把折现率调高，"
             "或把增长率调低到折现率以下，再重新计算。")

    # ---------- B. 自定义现金流序列 ----------
    sub("B. 自定义现金流序列（任意期数、任意金额）")
    if interactive:
        print("    输入规则：逗号分隔的净现金流；投入填负数、收回填正数；第 0 期代表期初投入。")
        raw = input("    各期净现金流（直接回车采用默认 "
                    f"{','.join(f'{c:g}' for c in _DEFAULT_CASHFLOWS)}）: ").strip()
        try:
            cfs = _parse_cashflows(raw) if raw else list(_DEFAULT_CASHFLOWS)
        except core.FinanceError as ex:
            err(f"{ex} 已改用默认序列继续演示。")
            cfs = list(_DEFAULT_CASHFLOWS)
        r2 = ask("每期折现率 r（与序列同口径）", 0.08, "rate")
        ppy = int(ask("每年现金流发生次数（1=年/2=半年/4=季/12=月）", 1, None))
    else:
        raw_cfs = p.get("cashflows", _DEFAULT_CASHFLOWS)
        # 非交互入口的现金流既可能是列表（CLI 预设 / run_all），
        # 也可能是逗号分隔字符串（tkinter 界面把文本字段原样传来），两种都支持
        cfs = _parse_cashflows(raw_cfs) if isinstance(raw_cfs, str) else [float(c) for c in raw_cfs]
        r2 = p.get("cf_rate", 0.08)
        ppy = int(p.get("periods_per_year", 1))

    sched = core.fv_from_schedule(cfs, r2)
    npv_value = core.npv(r2, cfs)
    n_cf = len(cfs)
    per = _PERIOD_UNIT.get(ppy, "期")
    total_out = sum(c for c in cfs if c > 0)
    compound_factor = (1.0 + r2) ** (n_cf - 1)

    kv_table({"现金流期数": n_cf, "每年发生次数": ppy, "每期折现率": r2,
              "累计投入": sched["total_in"], "累计收回": total_out,
              "净现金流入(收回-投入)": total_out - sched["total_in"],
              "滚动终值 FV": sched["fv"], "折现净现值 NPV": npv_value},
             "自定义现金流计算结果",
             pct_keys={"每期折现率"},
             money_keys={"累计投入", "累计收回", "净现金流入(收回-投入)",
                         "滚动终值 FV", "折现净现值 NPV"})

    ear2 = (1.0 + r2) ** ppy - 1.0
    info(f"口径声明：以上按【期利率口径】计算——每期折现率 {r2:.4%}（每{per}一期）、共 {n_cf} 期，"
         f"折合年数 {n_cf / ppy:.2f} 年、折合年化有效利率 {ear2:.2%}。")
    info(f"滚动终值法：自第 0 期起逐期滚存 FV_t = FV_(t-1) × (1 + {r2:.4%}) + CF_t，"
         f"得期末终值 {sched['fv']:,.2f} 元——回答的是“各期现金流按 {r2:.4%} 再投资，期末能积累多少”。")
    info(f"折现法交叉验证：把同一序列按 {r2:.4%} 折现，净现值 NPV = {npv_value:,.2f} 元。"
         f"两个指标回答的问题不同：滚动终值看“期末积累”，NPV 看“今天的价值”。"
         f"序列含正负现金流（即投资型现金流）时，NPV 才是决策依据——"
         f"NPV {'> 0，说明该现金流的收益率高于要求回报率，投资可行' if npv_value > 0 else '≤ 0，说明收益未达到要求回报率，应谨慎'}。")
    info(f"内嵌校验（两法互验）：滚动终值应等于折现净现值按同一利率滚存到期末，即 "
         f"FV = NPV × (1 + r)^(期数-1) = {npv_value:,.2f} × {compound_factor:.6f} = "
         f"{npv_value * compound_factor:,.2f} 元，与实际输出 {sched['fv']:,.2f} 元一致"
         f"（差异 {abs(npv_value * compound_factor - sched['fv']):.6f} 元，为浮点误差）。"
         f"这说明“逐期滚动”与“整体折现”两条独立实现路径相互吻合。")

    fig_path = os.path.join(OUT_DIR, "11_自定义现金流序列.png")
    vz.plot_custom_cashflow(cfs, r2, fig_path, path=sched["path"], periods_per_year=ppy)
    ok(f"自定义现金流序列图已生成：{os.path.relpath(fig_path, BASE_DIR)}")

    return {"场景": "永续年金与自定义现金流",
            "每期PMT": pmt, "折现率r": r, "增长率g": g,
            "永续年金现值PVP": pvp if pvp is not None else "",
            "现金流期数": n_cf, "每年发生次数": ppy, "每期折现率": r2,
            "累计投入": sched["total_in"], "累计收回": total_out,
            "滚动终值FV": sched["fv"], "折现净现值NPV": npv_value,
            "_cashflow_df": pd.DataFrame({"期数": list(range(n_cf)),
                                          "净现金流": cfs,
                                          "滚动终值": sched["path"]})}


# =============================================================================
# 一键运行全部场景
# =============================================================================
def run_all(interactive: bool = False) -> None:
    """按顺序执行全部示例场景，生成图表与 Excel 汇总报告。"""
    title("一键运行 · 全部示例场景")
    results = [
        scene_compound(False, {"pv": 100_000, "rate": 0.06, "years": 10, "freq": 1, "mode": "1"}),
        scene_annuity(False, {"pmt": 10_000, "rate": 0.08, "periods": 10, "due": False,
                              "periods_per_year": 1}),
        scene_debt(False, {"face": 1000, "coupon": 0.06, "years": 5, "price": 950, "tax": 0.25}),
        scene_equity(False, {"rf": 0.025, "beta": 1.2, "rm": 0.095, "use_ddm": True,
                             "d1": 0.5, "p0": 12.0, "g": 0.04}),
        scene_wacc(False, {"face": 1000, "coupon": 0.06, "years": 5, "price": 980, "tax": 0.25,
                           "rf": 0.025, "beta": 1.15, "rm": 0.095,
                           "equity": 6000, "debt": 4000, "pref": 0.0}),
        scene_loan(False, {"principal": 1_000_000, "rate": 0.042, "years": 20, "freq": 12}),
        scene_cashflow(False, {"pmt": 50_000, "rate": 0.08, "growth": 0.02,
                               "cashflows": _DEFAULT_CASHFLOWS, "cf_rate": 0.08,
                               "periods_per_year": 1}),
    ]

    # ---- 汇总 Excel ----
    xlsx = os.path.join(OUT_DIR, "09_计算结果汇总.xlsx")
    rp.export_summary(results, xlsx)
    ok(f"计算结果汇总 Excel 已生成：{os.path.relpath(xlsx, BASE_DIR)}")

    # ---- 汇总 Markdown 报告（供撰写作业报告直接引用） ----
    md = os.path.join(OUT_DIR, "计算结果汇总.md")
    rp.export_markdown(results, md)
    ok(f"汇总 Markdown 已生成：{os.path.relpath(md, BASE_DIR)}")

    sub("全部场景执行完毕")
    info(f"所有图表与数据文件位于：{OUT_DIR}")
    for f in sorted(os.listdir(OUT_DIR)):
        info(f"  · {f}")


# =============================================================================
# 交互菜单
# =============================================================================
MENU = """
   ┌──────────────────────────────────────────────────────────┐
   │       货币时间价值与资本成本计算器  v1.0                │
   │       Time Value of Money & Cost of Capital Calculator   │
   ├──────────────────────────────────────────────────────────┤
   │  第 3 章 货币时间价值                                    │
   │    1. 复利终值 / 现值计算                                │
   │    2. 年金终值 / 现值计算（普通年金 / 预付年金）          │
   │    6. 贷款还款计划（等额本息，年金公式应用）              │
   │    7. 永续年金现值 / 自定义现金流序列                     │
   │  第 4 章 资本成本                                        │
   │    3. 债权资本成本（含税盾效应）                          │
   │    4. 股权资本成本（CAPM / DDM）                          │
   │    5. 加权平均资本成本 WACC（含敏感性分析）               │
   │  其他                                                    │
   │    8. 一键运行全部示例场景（生成全部图表与 Excel）        │
   │    0. 退出程序                                            │
   └──────────────────────────────────────────────────────────┘
"""


def interactive_menu() -> None:
    """交互式主菜单循环。"""
    handlers = {
        "1": scene_compound, "2": scene_annuity, "3": scene_debt,
        "4": scene_equity, "5": scene_wacc, "6": scene_loan,
        "7": scene_cashflow,
    }
    while True:
        print(_c(MENU, "36"))
        try:
            choice = input("   请选择功能编号: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n    已退出，感谢使用。")
            return

        if choice == "0":
            print("\n    已退出，感谢使用。")
            return
        if choice == "8":
            run_all(False)
            continue
        h = handlers.get(choice)
        if not h:
            err("无效的功能编号，请输入 0~8 之间的数字。")
            continue
        try:
            h(True)
        except KeyboardInterrupt:
            warn("已取消当前计算，返回主菜单。")
        except core.FinanceError as ex:
            err(f"计算失败：{ex}")
        except Exception:                                   # noqa: BLE001
            err("程序发生未预期错误，详细信息如下：")
            traceback.print_exc()
        print()


# =============================================================================
# 入口
# =============================================================================
def main() -> None:
    parser = argparse.ArgumentParser(
        description="货币时间价值与资本成本计算器（Python + Matplotlib）")
    parser.add_argument("--demo", action="store_true", help="一键运行全部示例场景，输出图表与 Excel")
    parser.add_argument("--scene", type=int, choices=[1, 2, 3, 4, 5, 6, 7],
                        help="直接运行指定编号的场景（非交互）")
    parser.add_argument("--gui", action="store_true",
                        help="启动 tkinter 图形界面（不改变上述命令行功能）")
    args = parser.parse_args()

    # ---- 图形界面：实现全部在 gui.py，本文件只负责转发 ----
    if args.gui:
        try:
            import gui
        except Exception as exc:                        # noqa: BLE001
            err(f"图形界面启动失败：{type(exc).__name__}: {exc}")
            err("请先安装依赖：python -m pip install pillow（tkinter 随 Python 自带）")
            return
        gui.launch()
        return

    banner = r"""
   ╔══════════════════════════════════════════════════════════╗
   ║   货币时间价值与资本成本计算器                            ║
   ║   Time Value of Money & Cost of Capital Calculator       ║
   ║   第 3 章 复利/年金  ·  第 4 章 债务/股权/WACC            ║
   ╚══════════════════════════════════════════════════════════╝
"""
    print(_c(banner, "36"))

    if args.demo:
        run_all(False)
        return
    if args.scene:
        {"1": scene_compound, "2": scene_annuity, "3": scene_debt,
         "4": scene_equity, "5": scene_wacc, "6": scene_loan,
         "7": scene_cashflow}[str(args.scene)](False)
        return
    interactive_menu()


if __name__ == "__main__":
    main()
