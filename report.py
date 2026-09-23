#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
report.py —— 计算结果导出引擎（Excel 汇总 / Markdown 报告）
================================================================

职责
----
把各场景的"扁平摘要 dict"（main.py 中 scene_* 函数的返回值）导出为可直接
引用的交付物：

    09_计算结果汇总.xlsx    export_summary()        场景汇总 + 附加明细工作表
    计算结果汇总.md          export_markdown()       供作业报告直接引用
    07_贷款还款计划.xlsx    export_loan_schedule()  等额本息还款计划表
    10_测试用例明细.xlsx    export_casebook()       测试用例集明细

数值口径
--------
* 字段名命中"利率 / 成本 / 比重 / 权重 / 税率 / 溢价 / 收益率 / 增长率 / EAR /
  WACC"→ 按百分比渲染（保留 4 位小数）；
* 其余数值 → 千分位、保留 4 位小数；布尔值 → 是 / 否。
* 该口径与 visualize.py 的图表刻度、main.py 的终端输出一致，保证"屏=图=表"。

内部字段约定
------------
摘要 dict 中以 "_" 开头的键（如 _wacc_dict、_scenarios、_schedule_df）不写入
汇总表，而会被自动提取为附加工作表，便于审计与二次分析：

    _wacc_dict     dict                 → 「WACC明细」
    _scenarios     list[dict]           → 「资本结构模拟」
    _schedule_df   pandas.DataFrame     → 「还款计划前24期」
    _cashflow_df   pandas.DataFrame     → 「自定义现金流滚动路径」

调用方可直接用 extra_sheets 参数显式提供附加表；两种来源会合并并按表名去重。
"""

from __future__ import annotations

import os
from typing import Any, Dict, List, Sequence, Tuple

import pandas as pd

# -----------------------------------------------------------------------------
# 一、格式化口径
# -----------------------------------------------------------------------------
#: 命中这些关键词的字段按百分比渲染，其余按金额/数值渲染
_PCT_TOKENS = ("利率", "成本", "比重", "权重", "税率", "溢价",
               "收益率", "增长率", "EAR", "WACC")


def fmt_is_pct_key(key: str) -> bool:
    """按字段名语义判断是否应以百分比渲染。"""
    ku = str(key).upper()
    return any(t.upper() in ku for t in _PCT_TOKENS)


def fmt_summary_value(key: str, value: Any) -> str:
    """按字段名语义格式化摘要值：利率/成本/比重类 → 百分比；数值 → 千分位 4 位。"""
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


def _ensure_dir(path: str) -> str:
    """确保目标文件所在目录存在，返回原路径（与 visualize.py 同名工具函数一致）。"""
    d = os.path.dirname(os.path.abspath(path))
    if d:
        os.makedirs(d, exist_ok=True)
    return path


# -----------------------------------------------------------------------------
# 二、附件工作表：从摘要 dict 的内部字段自动提取
# -----------------------------------------------------------------------------
#: (内部字段名, 目标工作表名, 构造方式)
_SHEET_SOURCES: Tuple[Tuple[str, str], ...] = (
    ("_wacc_dict", "WACC明细"),
    ("_scenarios", "资本结构模拟"),
    ("_schedule_df", "还款计划前24期"),
    ("_cashflow_df", "自定义现金流滚动路径"),
)


def _to_frame(field: str, value: Any) -> "pd.DataFrame | None":
    """把内部字段转换为二维表；无法转换时返回 None。"""
    if isinstance(value, pd.DataFrame):
        return value
    if field == "_wacc_dict" and isinstance(value, dict):
        # core.wacc() 的返回字典：过滤内部键后转为「指标 | 数值」两列
        items = [(k, v) for k, v in value.items() if not str(k).startswith("_")]
        return pd.DataFrame({"指标": [k for k, _ in items],
                             "数值": [v for _, v in items]})
    if isinstance(value, (list, tuple)):
        rows = [r for r in value if isinstance(r, dict)]
        if rows:
            return pd.DataFrame(rows)
    return None


def _auto_extra_sheets(results: Sequence[Dict[str, Any]],
                       used_names: Sequence[str] = ()) -> List[Tuple[str, Any]]:
    """扫描各场景摘要 dict，把内部字段（_wacc_dict / _scenarios / _schedule_df）
    提取为附加工作表；重名时追加序号，避免 ExcelWriter 报错。"""
    sheets: List[Tuple[str, Any]] = []
    used = set(used_names)
    for r in results:
        if not isinstance(r, dict):
            continue
        for field, base_name in _SHEET_SOURCES:
            value = r.get(field)
            if value is None:
                continue
            df = _to_frame(field, value)
            if df is None or len(df) == 0:
                continue
            name, idx = base_name, 2
            while name in used:
                name, idx = f"{base_name}{idx}", idx + 1
            used.add(name)
            sheets.append((name, df))
    return sheets


# -----------------------------------------------------------------------------
# 三、导出函数
# -----------------------------------------------------------------------------
def export_summary(results: List[Dict[str, Any]], path: str,
                   extra_sheets: Sequence[Tuple[str, Any]] = ()) -> str:
    """
    导出计算结果汇总 Excel（09_计算结果汇总.xlsx）。

    参数
    ----
    results      : 各场景的扁平摘要 dict 列表（第一列通常为"场景"）
    extra_sheets : 附加工作表 [(sheet_name, DataFrame), ...]，如资本结构优化表。
                   未显式提供时，自动从摘要 dict 的内部字段提取（见模块说明）。
    """
    _ensure_dir(path)
    rows = [_sanitize_row(r) for r in results if isinstance(r, dict)]

    # 列顺序：各场景出现过的键依次排列，"场景"始终置顶
    cols: List[str] = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    if "场景" in cols:
        cols.remove("场景")
        cols.insert(0, "场景")

    df = pd.DataFrame(rows, columns=cols)

    # 附件表：显式提供的在前，自动提取的补充在后，同名去重
    sheets: List[Tuple[str, Any]] = []
    explicit_names: List[str] = []
    for name, extra in extra_sheets:
        if extra is None:
            continue
        sheet_name = str(name)[:31]
        explicit_names.append(sheet_name)
        sheets.append((sheet_name, extra))
    sheets += _auto_extra_sheets(results, used_names=explicit_names)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="场景汇总", index=False)
        for name, extra in sheets:
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
        for k, val in row.items():
            lines.append(f"| {k} | {fmt_summary_value(k, val)} |")
        lines.append("")
    lines += [
        "---",
        "",
        f"> 说明：本文件由 `{os.path.basename(__file__)}` 自动生成，"
        "数值口径为“利率=小数、金额=元”。",
        "> 完整图表见输出目录下的 PNG 文件。",
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
    """
    _ensure_dir(path)
    overview = [{"序号": i, "用例名称": rec.get("名称", ""), "结论": rec.get("结论", "")}
                for i, rec in enumerate(records, 1)]
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(overview).to_excel(writer, sheet_name="用例索引", index=False)
        for i, rec in enumerate(records, 1):
            body: List[Dict[str, Any]] = []
            for k, val in (rec.get("参数") or {}).items():
                body.append({"项目": "参数", "内容": f"{k} = {val}"})
            for item in (rec.get("过程") or []):
                if isinstance(item, (tuple, list)) and len(item) >= 2:
                    body.append({"项目": "计算步骤", "内容": f"{item[0]} → {item[1]}"})
                else:
                    body.append({"项目": "计算步骤", "内容": str(item)})
            body.append({"项目": "结论", "内容": rec.get("结论", "")})
            pd.DataFrame(body).to_excel(writer, sheet_name=f"用例{i}"[:31], index=False)
    return path
