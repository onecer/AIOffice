# fixtures/

测试夹具与黄金脚本（golden scripts）。Fixtures and golden scripts for AIOffice.

## 目录结构 / Layout

```
fixtures/
  scripts/         黄金脚本：JSON 形式的 CLI argv 数组 + 信封断言
    run_goldens.py 跨平台执行器（仅用 Python 标准库）
  manual-check/    供真人用 Word/Excel/PowerPoint 打开核验的产物（CI 不消费）
```

## 黄金脚本如何工作 / How golden scripts work

每个 `scripts/*.json` 是一串 CLI 调用。每一步是一个 argv 数组，执行器会自动
追加 `--workspace <临时目录> --json`，然后断言 stdout 恰好是一个 JSON 信封：

```json
{
  "argv": ["create", "golden-out/hello.docx", "--title", "Golden DOCX"],
  "expect": { "ok": true, "dataContains": "可选的子串断言", "errorCode": "仅当 ok=false" }
}
```

本地运行 / Run locally:

```bash
python3 fixtures/scripts/run_goldens.py
# 或指定已发布的二进制 / or against a published binary:
AIOFFICE_CMD=/path/to/aioffice python3 fixtures/scripts/run_goldens.py
```

CI 在 macOS 与 Windows 上跑同一批脚本（见 `.github/workflows/ci.yml`）。
三个脚本即三条端到端主路径：docx 创建+编辑+读取、xlsx 公式+取值、
pptx 加页+加形状+渲染 SVG。每个脚本都以 `validate` 收尾——这是
"OpenXmlValidator 0 错误" 质量门的端到端镜像。

## manual-check/

格式包的测试应把代表性产物（创建出的 docx/xlsx/pptx）额外写一份到
`fixtures/manual-check/`，由真人用真实 Office 打开确认。CI 无法自动化这一步，
OpenXmlValidator 只是代理指标。

M7（v0.8.0）三件套是**已审计 + 已修复**的干净产物，请用真实 Office 打开核验：

- `report.docx` —— 文档标题 "Quarterly Overview"、自定义属性 Project=Aurora / Reviewed=true
  （文件 → 信息 → 属性里可见）、图片带 alt 文本、表格首行标记为重复表头、一个值为 Final 的
  下拉**内容控件**（Status）。`aioffice audit report.docx` 应报告**零 findings**。
- `metrics.xlsx` —— 文档标题 "Sales Metrics"、命名单元格样式 **Currency-Red**（红色加粗货币格式，
  套在 B2:B4，Excel 单元格样式库里可见），`=SUM(B2:B3)` 应为 3600。`audit` 应报告**零 findings**。
- `deck.pptx` —— 幻灯片标题 "Q4 Review"、图片带 alt 文本、24pt 可读正文。`audit` 仅剩一条
  `a11y_reading_order`（info，非缺陷，提示图形在视觉上位于文本之上）。

`audit-demo/` 里是**故意做坏**的同名三件套（before-audit），供你亲自跑 `aioffice audit`
（及 `--fix`）观察真实 findings——详见 `audit-demo/README.md`。

M8（v0.9.0）对比与审阅产物，请用真实 Office 打开核验：

- `diff-base.docx` + `diff-new.docx` —— 一对用于 diff 的文档：new 相比 base 改了第 2 段文本、
  追加了一段、把标题样式从 Heading1 改成 Heading2。`aioffice diff diff-new.docx diff-base.docx`
  应报告 `{added:1, modified:2}`（1 个 text、1 个 style）。Word 里逐段对照即可肉眼复核。
- `report-captions.docx` —— 一个 **Figure 题注**（"Figure 1: Quarterly revenue trend (Q1–Q4)"，Caption 样式）
  + 一条指向它的**交叉引用**（"As shown in Figure 1"）。Word 打开后按 F9 刷新域，编号应自洽；
  右键交叉引用 → 更新域应跳转到题注。
- `dashboard-slicer.xlsx` —— Sheet1 上一个名为 **Sales** 的表 + 一个挂在 **Region** 列上的**切片器**
  （锚在 E2）。Excel 打开后切片器应可点选筛选表格行。
- `deck-hyperlinks.pptx` —— 第 1 页两个形状带**点击动作**：一个**跳转到第 2 页**（`#slide:2`）、
  一个**外链** https://aioffice.dev 。PowerPoint 放映态下点击应分别跳页 / 打开网址。

历史：M7（v0.8.0）三件套是已审计+已修复的干净 report.docx/metrics.xlsx/deck.pptx；M6（v0.7.0）三件套曾覆盖 docx 公式/RTL/分栏、xlsx Excel 表/大纲分组、pptx 母版/分节/幻灯片尺寸。

## 从 OfficeCLI 学习能力（而非语法）/ Lifting capability cases from OfficeCLI

`/tmp/office_research/OfficeCLI` 是能力对齐的上游参照（SKILL.md 为能力清单，
`src/officecli/Handlers/*` 为行为参考）。从它那里提取**测什么**，而不是**怎么写**：

1. 在 SKILL.md / Handlers 里找一个能力点（例如"合并单元格后读取范围值"）。
2. 记下它覆盖的边界情况（空单元格、共享字符串、跨 sheet 引用……）。
3. 用 **AIOffice 自己的命令面**把同一场景写成黄金脚本或 xunit 用例：
   我们的 `edit --ops`/寻址语法与 OfficeCLI 完全不同，这是有意为之的设计决策
   （能力对齐、语法自研），不要复刻它的命令、参数名或输出形状。
4. 绝不复制其代码。行为可以学习，实现必须自己写。

New capability cases should be lifted the same way: take the *scenario and its
edge cases* from the OfficeCLI reference, then express them in AIOffice's own
command surface. Never copy its command syntax or code.
