# 图形界面（GUI）开发说明与后续建议

> 适用项目：`financial_calculator` —— 货币时间价值与资本成本计算器
> 初版：在原有命令行程序之上新增 tkinter 图形界面，命令行功能全部保留。
> 最近更新（2026-09-23）：场景数由 6 个扩展为 **7 个**（新增"永续年金现值 / 自定义现金流序列"页）；
> 表单补齐付息频率、筹资费率、计息频率等参数；§3.2/§3.3/§3.6/§3.8 中列出的建议多数已落地。

---

## 一、本次改动清单

| 文件 | 状态 | 说明 |
| --- | --- | --- |
| `gui.py` | **新增** | 图形界面模块：执行适配层 + **7 个场景页** + 总览页 + 主窗口 |
| `financial_calculator_gui.pyw` | **新增** | 双击启动器（pythonw 方式，无黑色控制台窗口） |
| `main.py` | 修改 | ① 新增 `--gui` 参数；② 修复 `_USE_COLOR` 在 pythonw 下会崩的隐患；③ 新增场景 7 |
| `requirements.txt` | 修改 | 补充 `pillow`（界面图片缩放与截图功能所需）；版本改为区间声明 |
| `report.py` | 上一轮新增 | 导出引擎（本次未改动） |

场景页与命令行场景一一对应（字段数取自 `gui.SCENE_SPECS`）：

| 页 | 场景 | 字段数 | 备注 |
| --- | --- | --- | --- |
| 1 | 复利终值 / 现值 | 5 | — |
| 2 | 年金终值 / 现值 | 5 | 新增"每年收付次数"选择（1/2/4/12） |
| 3 | 债权资本成本 | 7 | 新增"每年付息次数"（1/2/4）与"筹资费率" |
| 4 | 股权资本成本 | 8 | — |
| 5 | 加权平均资本成本 WACC | 14 | 新增"每年付息次数"与"筹资费率" |
| 6 | 贷款还款计划 | 4 | — |
| 7 | 永续年金现值 / 自定义现金流 | 6 | **整页新增**，含 `kind="text"` 的现金流序列输入 |

改动后**没有删除任何命令行能力**：

```bash
python main.py                     # 交互式菜单（原有）
python main.py --demo              # 一键运行全部场景（原有）
python main.py --scene 3           # 直接运行指定场景（原有，范围 1~7）
python main.py --gui               # 图形界面（新增）
python gui.py                      # 图形界面（新增，等价写法）
双击 financial_calculator_gui.pyw  # 图形界面（新增，无控制台窗口）
python gui.py --screenshot 路径    # 生成界面截图，供报告插图（新增）
python gui.py --check              # 依赖与字体自检（新增）
```

---

## 二、这次是怎么接的：一个刻意选择的架构

### 2.1 核心决策——GUI 不重写计算，只重写"呈现"

项目原有的分层是：

```
core.py          计算引擎（纯函数）
validators.py    三层校验 + 业务解读
visualize.py     Matplotlib 图表
report.py        Excel / Markdown 导出
main.py          场景编排 + 命令行交互 + print 输出
```

新增 GUI 时有两条路：

1. **另写一套 GUI 版场景函数**（`run_compound`、`run_wacc`…，返回结构化结果）。
   代价：计算逻辑出现第二份实现，两个入口可能算出不同答案。
2. **复用 `main.py` 的场景函数，把它们的标准输出捕获到界面中渲染**。
   代价：界面依赖命令行输出的文本格式。

本次选了方案 2，实现方式是 `gui.capture()`：

```python
def capture(name, call, out_dir=None) -> RunOutput:
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):     # 捕获 CLI 的 print
        summary = call()
    text = strip_ansi(buffer.getvalue())                 # 去掉 ANSI 颜色码
    figures, files = _collect_outputs(out_dir, started)  # 按修改时间采集产物
    return RunOutput(name, text, figures, files, summary)
```

场景页调用时就是：

```python
self.app.submit(title, lambda: main.scene_wacc(False, params), self)
```

好处很直接：

* **单一事实来源**：命令行与图形界面的计算结果、校验提示、业务解读逐字一致，不存在"两套答案"。
* **零重复代码**：没有为 GUI 重写任何财务逻辑。
* **改一处、两处生效**：以后调整某个公式的解释文本，命令行和界面同步更新。

代价也要说清楚：

* 界面渲染依赖命令行输出的符号体系（`═` 标题、`▶` 小节、`✔/⚠/✘` 结果、`─` 表格）。目前靠 `classify_lines()` 做映射，如果 `main.py` 改了输出符号，`gui.py` 的渲染规则要同步调整。
* 只能整段呈现文本，做不到分节折叠、逐项高亮这类细粒度交互。

**如果后续要长期迭代这个项目，建议按下面 3.1 的结构化改造走**——那是这条路的自然终点。

### 2.2 线程模型

`main.py` 的场景函数是同步阻塞的（计算 + Matplotlib 绘图 + 写 Excel，一键运行约 5 秒）。
GUI 用「后台线程执行 + 主线程轮询队列」的方式，避免窗口假死：

```
主线程 submit()  →  加锁置忙、启动进度条  →  daemon 线程执行 capture()
主线程 after(80ms) 轮询队列  →  取到结果后刷新结果区与图表预览
```

同时用 `_busy` 标志把任务串行化：Matplotlib 的 pyplot 全局状态不是线程安全的，串行执行可以规避竞争。

### 2.3 中文表格对齐的坑

命令行输出用空格对齐表格，界面上必须用**中英文 2:1 等宽字体**才能保持对齐。
Consolas / Cascadia 这类西文等宽字体的中文会回退到雅黑，实测宽高比约 1.78:1，表格会逐行错位。
因此 `gui.resolve_fonts()` 按 `新宋体 → NSimSun → 宋体 → SimSun → Consolas` 的顺序探测，
结果区固定用第一个命中的字体。**不要为了"好看"把结果区换成 Consolas。**

---

## 三、后续开发建议（按优先级）

### 3.1 【高】让场景函数返回结构化结果，彻底解耦输出

现状：`main.scene_*()` 一边 `print` 一边返回一个扁平 `dict`，GUI 只能捕获 stdout。
目标：让场景层返回结构化对象，CLI 与 GUI 各自渲染自己需要的部分。

```python
# 建议新增到 core 或单独的 scene.py
@dataclass
class SceneResult:
    name: str                                            # 场景名
    summary: Dict[str, Any]                              # 扁平摘要（Excel / 状态栏）
    sections: List[Tuple[str, List[str]]]                # [(小节标题, 文本行)]
    figures: List[str] = field(default_factory=list)     # 生成的 PNG
    files: List[str] = field(default_factory=list)       # 生成的其他文件
    extra_sheets: List[Tuple[str, Any]] = field(default_factory=list)
```

改造步骤（可以逐个场景做，不必一次到位）：

1. 在 `main.py` 旁边加一个 `render_cli(result)`：把 `sections` 按现在的符号体系打印出来，保证命令行输出**逐字不变**。
2. 把 `scene_compound` 等函数内的 `kv_table(...)` / `info(...)` 调用改为往 `sections` 里追加文本行。
3. `main.py` 的每个场景调用处包一层 `render_cli(...)`。
4. `gui.py` 直接消费 `result.sections`，不再需要 `capture()` 和 `classify_lines()`。

收益：GUI 不再依赖 stdout 文本格式；可以按小节分块展示、加折叠、给每张图配标题；也方便以后导出成 Word/PDF 报告。

> 顺带一提：历史上的单文件版 `financial_calculator_optimized.pyw` 曾实现过这套 `SceneResult` 结构。
> 该文件已在 2026-09-23 的清理中删除（如需查阅，可执行 `git checkout HEAD -- financial_calculator_optimized.pyw` 恢复），
> 其 `sections / extra_sheets` 设计思路仍可作为改造参考，避免重新设计。

### 3.2 【高】统一输出目录，消除 `outputs/` 与 `output/` 并存 —— ✅ 已确认无需处理

核查结论：项目当前**只存在 `outputs/` 一个输出目录**，旧版遗留的 `output/` 并不存在，因此两套入口不会写到不同地方。

`main.py` 中 `OUT_DIR = BASE_DIR / "outputs"` 即唯一事实来源，`gui.py` 统一读取 `main.OUT_DIR`。
`.gitignore` 已同时忽略 `outputs/` 与 `output/`，即使将来出现 `output/` 也不会误提交。

> 若后续确有需要，仍可按下方两种方案之一收敛；当前无此必要，不建议为"可能性"改代码。

```python
# 方案 A：沿用已存在的目录，避免产生两份产物（与 .pyw 版行为一致）
def resolve_out_dir() -> str:
    for name in ("outputs", "output"):
        candidate = os.path.join(BASE_DIR, name)
        if os.path.isdir(candidate):
            return candidate
    return os.path.join(BASE_DIR, "outputs")

OUT_DIR = resolve_out_dir()
```

```python
# 方案 B：直接统一到 output/
OUT_DIR = os.path.join(BASE_DIR, "output")
```

### 3.3 【中】把 `test_cases.py` 改造成可汇总的测试 —— ✅ 已完成

原问题：`check()` 断言失败直接 `raise AssertionError`，第一个失败就中断，且没有汇总。

已落地的改造：

* `check()` 改为**记录失败并继续执行**（维护 `_passed` / `_failed` 两个列表），单个用例失败不再中断整轮；
* 末尾输出汇总行 `✔ 通过 51 项 / 失败 0 项（共 51 项数值断言）`，并以 `sys.exit(1 if _failed else 0)` 返回**退出码**，便于 CI 与脚本判定成败；
* 新增 `expect_guard()` 辅助，专门断言"非法输入必须抛中文 `FinanceError`"，同时捕获"未受控异常"与"静默返回错误结果"两类失败模式；
* 新增 `case_boundary()` 边界回归组（11 项），并把用例总数从 7 组扩充为 8 组。

后续可选：迁移到 `pytest`（`pytest test_cases.py -q`），使 CI 与本地执行方式统一。当前不阻塞。

### 3.4 【中】参数方案与结果的"留存"能力

面向作业/汇报场景，这几项投入产出比很高：

* **参数方案保存/加载**：把表单值存成 JSON（`方案/房贷.json`），下次一键载入，报告里也能附上参数。
* **结果导出**：把当前场景的文本 + 图表导出为一份 Word 或 PDF（已有 `docx` 能力可直接用），省去手工截图排版。
* **批量敏感性分析**：勾选一个参数（如利率 3%~6%），一次跑多组并输出对照表/图。
* **历史记录**：内存中保留最近 20 次计算（时间、场景、关键结果），便于对比。

### 3.5 【中】交互体验细节

* 输入框失焦即校验（`<FocusOut>` 绑定 `vd.validate`），错误当场标红，不必等点击计算。
* `<Return>` 回车即计算，符合直觉。
* 结果区已支持横向滚动（`wrap="none"` + 双向滚动条），如果以后改用结构化渲染，可以按小节分栏。
* 深色/浅色主题切换：`UI` 字典已经是集中配置，加一套浅色配色切换即可。
* 界面文字统一：目前场景名在页面标题用"｜"、在命令行输出里用"|"，如需完全一致可统一为一种。

### 3.6 【中】工程质量 —— 前两项已完成

* **依赖固定**：✅ 已完成。`requirements.txt` 改为"下限 + 下一主版本上限"的区间声明（`numpy>=2.0,<3`、`pandas>=2.0,<3`、`matplotlib>=3.7,<4`、`openpyxl>=3.1,<4`、`pillow>=10.0,<12`），并注明实测通过的环境。
  为什么不精确锁定（`==`）：本项目是教学交付物，需要在只有较新版本的机器上也能装上；区间写法既避免环境漂移，又不至于装不上。真正的精确复现应交给 `pip freeze > requirements.lock`。
* **`.gitignore`**：✅ 已完成，并修正了一处会造成交付事故的写法。原 `.gitignore` 用 `*.png` / `*.pdf` 全局通配，会把 `docs/figures/` 的报告插图和 `docs/开发报告.pdf` 一并忽略——本地看不出问题，但推送到远程后报告里的插图全部 404。现改为**按目录忽略**（`outputs/`、`output/`、`gui_临时截图/`），并显式保留 `!docs/figures/*.png`、`!docs/excel/*.xlsx`、`!docs/*.pdf`。
* **检查工具**：`ruff`（lint）+ `mypy`（类型）。目前代码已带类型注解，接入成本低——待做。
* **日志**：把散落的 `print` 逐步换成 `logging`，GUI 可以加一个"运行日志"面板，出问题时用户能自助排查——待做。

### 3.7 【低】打包发布

如果要做成"拷给别人双击就能用"的版本：

```bash
pip install pyinstaller
pyinstaller --noconfirm --windowed --name 财务计算器 ^
            --add-data "outputs;outputs" gui.py
```

注意三点：

1. **中文字体**：Matplotlib 依赖系统字体，打包后在没有雅黑/黑体的机器上中文会变方块，需要随包附带一个 CJK 字体并注册。
2. **路径**：打包后 `__file__` 指向解压目录，读写输出目录要用 `sys._MEIPASS` 判断或改用用户目录（`%USERPROFILE%\财务计算器输出`）。
3. **`.pyw` 与 `pythonw`**：打包时用 `--windowed` 等价于无控制台；调试阶段先不要加，保留控制台才能看到报错。

### 3.8 【低】清理历史文件 —— ✅ 已完成

原状况：项目里并存两套实现——

* `financial_calculator.pyw`（约 146 KB）与 `financial_calculator_optimized.pyw`（约 152 KB）：单文件整合版，自带 GUI；
* 模块版：`core.py / validators.py / visualize.py / report.py / main.py / gui.py`。

已按建议以**模块版为准**完成收敛：

* `financial_calculator.pyw` 已移入 `legacy/`（该目录内含 `README.md`，写明"历史版本，仅供对比，不参与测试/打包/报告引用"）；
* `financial_calculator_optimized.pyw` 已删除，可从历史提交恢复（`git checkout HEAD -- financial_calculator_optimized.pyw`）；
* 项目根目录不再直接暴露这两个 `.pyw`，"修改公式时只需改模块版"这一约定因此可以被机械地执行。

---

## 四、启动与验证速查

```bash
# 环境自检（依赖 / 字体 / 输出目录）
python gui.py --check

# 启动图形界面
python gui.py                 # 等价于双击 financial_calculator_gui.pyw
python main.py --gui          # 等价写法

# 命令行回归（确认 GUI 改动没影响原功能）
python main.py --demo
python main.py --scene 3      # 场景范围 1~7

# 生成界面截图（报告插图用）
python gui.py --screenshot outputs/gui_场景页预览.png
```

界面截图共两张：`gui_总览页.png`（总览页）与 `gui_场景页预览.png`（场景页），
已归档到 `docs/figures/`，是《开发报告》第八章 8.3 节的插图来源，可直接放进作业报告。

---

## 五、踩坑清单（下次改代码前先看这一条）

| 坑 | 现象 | 处理 |
| --- | --- | --- |
| `pythonw` 下 `sys.stdout` 为 `None` | `.pyw` 一启动就崩在 `sys.stdout.isatty()` | 用 `getattr(sys.stdout, "isatty", ...)` 包裹（已在 `main.py` 修复） |
| 模块内入口函数与导入模块同名 | `def main()` 遮蔽 `import main`，所有 `main.OUT_DIR` 报 `AttributeError` | 入口函数不要叫 `main`（本模块叫 `entry`） |
| 结果区用 Consolas | 命令行表格在界面里逐行错位 | 结果区必须用中文 2:1 等宽字体（新宋体等） |
| ANSI 颜色码残留 | 界面里出现 `[36m` 这类乱码 | `capture()` 中 `strip_ansi()` 剥离（`main._USE_COLOR` 也置 False） |
| DPI 125% 下界面模糊 | 文字发虚、控件偏小 | `enable_dpi_awareness()`（启动时调用，已实现） |
| Matplotlib 线程安全 | 并发绘图偶发崩溃/图片串味 | 保持 `_busy` 串行执行；将来要并行请改用进程池 |
| 路径含全角括号 `（8）` | 脚本里外引号/编码处理容易出错 | 全程用 `os.path.join`，不要在命令里手写路径字符串 |
