#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main.py —— 货币时间价值与资本成本计算器（命令行交互主程序）
==============================================================

运行方式
--------
    python3 main.py              # 交互式菜单模式
    python3 main.py --demo       # 一键运行全部示例并输出图表与 Excel
    python3 main.py --scene 1    # 直接运行指定场景（1~5）

功能菜单
--------
    1. 复利终值 / 现值计算              (第 3 章)
    2. 年金终值 / 现值计算              (第 3 章)
    3. 债权资本成本计算（含税盾效应）    (第 4 章)
    4. 股权资本成本计算（CAPM / DDM）    (第 4 章)
    5. 加权平均资本成本 WACC 计算        (第 4 章)
    6. 资本结构优化对比（多方案）        (第 4 章)
    7. 贷款还款计划（等额本息）          (第 3 章应用)
    8. 一键运行全部示例场景（生成图表 + Excel）
    0. 退出

设计说明
--------
* 交互层（input/output）与计算层（core.py）完全解耦，便于测试与二次开发。
* 每次计算前都经过 validators 三层校验，非法输入会被拦截并给出通俗提示。
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

# -----------------------------------------------------------------------------
# 终端输出美化（ANSI 颜色，Windows 终端若不支持可自动降级）
# -----------------------------------------------------------------------------
_USE_COLOR = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None


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


def ask(prompt: str, default: Optional[float] = None, rule: Optional[str] = None) -> float:
    """
    带校验的输入函数：循环直至通过 validators 校验。
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
             pct_keys={"年名义利率", "实际年利率 EAR"}, money_keys={"现值 PV", "复利终值 FV", "终值 FV", "复利现值 PV", "利息收益"} if False else None)

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
def scene_annuity(interactive: bool = True,
                  preset: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """普通年金 / 预付年金的终值与现值 + 现金流时间轴 + 折现贡献分解。"""
    title("场景 2 | 年金终值与现值计算（第 3 章）")
    if interactive:
        print("    提示：期利率 = 年利率 / 每年收付次数，请保持口径一致。")
        pmt = ask("每期现金流 PMT（元）", 10000, "positive_amount")
        r = ask("每期利率 r", 0.08, "rate")
        n = ask("期数 n", 10, "periods")
        due = (input("    收付方式（1=期末/普通年金，2=期初/预付年金）[1]: ").strip() or "1") == "2"
    else:
        p = preset or {}
        pmt, r, n = p.get("pmt", 10000), p.get("rate", 0.08), p.get("periods", 10)
        due = p.get("due", False)

    pva = core.pva_due(pmt, r, n) if due else core.pva_ordinary(pmt, r, n)
    fva = core.fva_due(pmt, r, n) if due else core.fva_ordinary(pmt, r, n)
    total = pmt * n

    kv = {"每期现金流 PMT": pmt, "每期利率": r, "期数": n,
          "收付方式": "预付年金（期初）" if due else "普通年金（期末）",
          "年金现值 PVA": pva, "年金终值 FVA": fva,
          "累计现金流总额": total, "现值/累计比": pva / total,
          "终值/累计比": fva / total}
    kv_table(kv, "计算结果", pct_keys={"每期利率", "现值/累计比", "终值/累计比"},
             money_keys={"每期现金流 PMT", "年金现值 PVA", "年金终值 FVA", "累计现金流总额"})

    sub("业务解读")
    for line in vd.interpret_tvm("年金现值" if not due else "预付年金现值",
                                 {"rate": r, "periods": n, "amount": pmt}, pva):
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

    return {"场景": "年金终值/现值", "每期PMT": pmt, "期利率": r, "期数": n,
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
        price = ask("发行价 / 净筹资额（元）", 950, "positive_amount")
        tax = ask("所得税税率", 0.25, "tax_rate")
    else:
        p = preset or {}
        fv_, cpn = p.get("face", 1000), p.get("coupon", 0.06)
        yrs, price, tax = p.get("years", 5), p.get("price", 950), p.get("tax", 0.25)

    d = core.cost_of_debt(fv_, cpn, int(yrs), price, tax)
    d_approx = core.cost_of_debt_approx(fv_, cpn, int(yrs), price, tax)

    kv = {"债券面值": fv_, "票面利率": cpn, "期限(年)": yrs, "发行价": price,
          "所得税税率": tax, "年利息支出": d["年利息支出"],
          "税前资本成本(YTM法)": d["税前年化资本成本"],
          "税后资本成本": d["税后年化资本成本"],
          "税盾节省(百分点)": d["税盾节省(百分点)"] / 100,
          "税前资本成本(近似公式法)": d_approx["税前年化资本成本(近似)"],
          "两法差异": d["税前年化资本成本"] - d_approx["税前年化资本成本(近似)"]}
    kv_table(kv, "计算结果",
             pct_keys={"票面利率", "所得税税率", "税前资本成本(YTM法)", "税后资本成本",
                       "税盾节省(百分点)", "税前资本成本(近似公式法)", "两法差异"},
             money_keys={"债券面值", "发行价", "年利息支出"})

    sub("方法学说明与业务解读")
    info("方法一（主算法）IRR/YTM 法：令债券未来本息现金流现值等于发行价的折现率，"
         "即债权人要求的到期收益率，也是企业真实承担的税前债务成本。")
    info("方法二（校验法）教材近似公式：分母取面值与发行价均值。两法结果高度接近，"
         f"差异仅 {abs(d['税前年化资本成本'] - d_approx['税前年化资本成本(近似)']):.4%}，互为验证。")
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
    vz.plot_debt_cost_curve(fig_path, fv_, cpn, int(yrs), tax_rate=tax)
    ok(f"债权资本成本曲线已生成：{os.path.relpath(fig_path, BASE_DIR)}")

    return {"场景": "债权资本成本", "面值": fv_, "票面利率": cpn, "期限(年)": yrs,
            "发行价": price, "税率": tax, "税前成本": d["税前年化资本成本"],
            "税后成本": d["税后年化资本成本"]}


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
        price = ask("发行价 / 净筹资额（元）", 980, "positive_amount")
        tax = ask("所得税税率", 0.25, "tax_rate")
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
        rf, beta, rm = p.get("rf", 0.025), p.get("beta", 1.15), p.get("rm", 0.095)
        e_val, d_val, p_val = p.get("equity", 6000), p.get("debt", 4000), p.get("pref", 0.0)
        kp = p.get("kp", 0.0)

    # ---- 债务成本 ----
    d = core.cost_of_debt(fv_, cpn, int(yrs), price, tax)
    kd_after = d["税后年化资本成本"]
    # ---- 股权成本 ----
    e = core.capm_required_return(rf, beta, rm)
    ks = e["股权资本成本"]
    # ---- 权重口径校验 ----
    w_res = vd.check_wacc_weights(d_val / (e_val + d_val + p_val),
                                  e_val / (e_val + d_val + p_val),
                                  p_val / (e_val + d_val + p_val))
    for wmsg in w_res.warnings:
        warn(wmsg)
    # ---- WACC ----
    w = core.wacc(e_val, d_val, ks, kd_after, p_val, kp)
    w["_税前债权成本"] = d["税前年化资本成本"]

    kv_table({
        "税前债务成本 Kd(税前)": d["税前年化资本成本"],
        "税后债务成本 Kd(税后)": kd_after,
        "股权成本 Ks(CAPM)": ks,
        "股权市场价值 E": e_val, "债务市场价值 D": d_val, "优先股 P": p_val,
        "股权权重 wE": w["股权比重"], "债务权重 wD": w["债权比重"],
        "WACC": w["WACC"],
    }, "WACC 汇总结果",
        pct_keys={"税前债务成本 Kd(税前)", "税后债务成本 Kd(税后)", "股权成本 Ks(CAPM)",
                  "股权权重 wE", "债务权重 wD", "WACC"},
        money_keys={"股权市场价值 E", "债务市场价值 D", "优先股 P"})

    sub("业务解读")
    for line in vd.interpret_wacc(w, industry_avg=0.10, roic=0.118, rf=rf):
        info(line)

    sub("资本结构优化：不同负债水平下的 WACC 模拟")
    tax_shield_note = "假设负债上升会小幅推高财务风险，从而抬升 Kd 与 Ks"
    info(tax_shield_note + "；作为对照进行情景测算。")
    scenarios = simulate_capital_structure(e_val, d_val, ks, d["税前年化资本成本"], tax)
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
            "股权成本Ks": ks, "股权权重": w["股权比重"], "债务权重": w["债权比重"], "WACC": w["WACC"],
            "_wacc_dict": w, "_scenarios": scenarios}


def simulate_capital_structure(
    equity: float, debt: float, ks_base: float, kd_pre_base: float, tax: float,
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
    """
    total = equity + debt
    wd0 = debt / total
    rows = []
    for wd in np.linspace(0.10, 0.80, 8):
        delta = wd - wd0
        kd_pre = kd_pre_base * (1 + 0.20 * delta)
        ks = ks_base * (1 + 1.05 * delta)

        # 高杠杆区的财务困境成本加速项
        if wd > 0.55:
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
    info(f"每月需还款 {s['每期还款额']:,.2f} 元，20 年累计还款 {s['总还款额']:,.2f} 元，"
         f"其中利息高达 {s['总利息']:,.2f} 元，相当于本金的 {s['总利息'] / principal:.1%}。")
    info(f"校验逻辑：把每月还款额按 {rate / freq:.4%} 的月利率折现 {years * freq} 期，"
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
# 一键运行全部场景
# =============================================================================
def run_all(interactive: bool = False) -> None:
    """按顺序执行全部示例场景，生成图表与 Excel 汇总报告。"""
    title("一键运行 · 全部示例场景")
    results = [
        scene_compound(False, {"pv": 100_000, "rate": 0.06, "years": 10, "freq": 1, "mode": "1"}),
        scene_annuity(False, {"pmt": 10_000, "rate": 0.08, "periods": 10, "due": False}),
        scene_debt(False, {"face": 1000, "coupon": 0.06, "years": 5, "price": 950, "tax": 0.25}),
        scene_equity(False, {"rf": 0.025, "beta": 1.2, "rm": 0.095, "use_ddm": True,
                             "d1": 0.5, "p0": 12.0, "g": 0.04}),
        scene_wacc(False, {"face": 1000, "coupon": 0.06, "years": 5, "price": 980, "tax": 0.25,
                           "rf": 0.025, "beta": 1.15, "rm": 0.095,
                           "equity": 6000, "debt": 4000, "pref": 0.0}),
        scene_loan(False, {"principal": 1_000_000, "rate": 0.042, "years": 20, "freq": 12}),
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
    parser.add_argument("--scene", type=int, choices=[1, 2, 3, 4, 5, 6],
                        help="直接运行指定编号的场景（非交互）")
    args = parser.parse_args()

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
         "4": scene_equity, "5": scene_wacc, "6": scene_loan}[str(args.scene)](False)
        return
    interactive_menu()


if __name__ == "__main__":
    main()
