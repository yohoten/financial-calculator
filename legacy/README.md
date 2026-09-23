# legacy —— 历史单文件版（仅供对照，不再维护）

本目录存放项目早期"单文件整合版"的实现：把计算、校验、绘图、界面全部写在
一个 `.pyw` 文件里，双击即可运行，不依赖本项目的模块划分。

| 文件 | 说明 |
| --- | --- |
| `financial_calculator.pyw` | 早期单文件整合版（自带 tkinter 界面），本目录保留的最后一份 |

> 另有一份 `financial_calculator_optimized.pyw`（152 KB）曾与上面这份并存，
> 已从工作区删除；如需找回，可用 `git checkout HEAD -- financial_calculator_optimized.pyw`
> 从历史提交中恢复。

## 为什么保留

自 v1.0 起项目以**模块版**为准：

```
core.py / validators.py / visualize.py / report.py / main.py / gui.py
```

单文件版仍有两处参考价值：

1. 它已经实现过一套 `SceneResult` 结构化结果（`sections` / `extra_sheets`），
   模块版后续若要把"场景函数返回结构化对象、CLI 与 GUI 各自渲染"落地，
   可以直接参考这里的设计，不必从零设计（详见 `docs/GUI开发说明与建议.md` §3.1）。
2. 它展示了另一种输出目录策略（"优先沿用已存在的 `outputs/` 或 `output/`"），
   对统一输出目录的取舍有参考意义。

## 注意事项

- **修改公式时只需改模块版，不要在此目录同步修改。**
  两套实现同时存在时，"同一个公式在两个地方各自演化"是这类项目最常见的维护事故。
- 本目录的代码**不参与**测试、打包与报告引用；`test_cases.py`、`outputs/` 与
  `docs/` 下的全部结果均来自模块版。
