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
