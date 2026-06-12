# AIOffice 能力账本（Capability Ledger）

> 本账本只记录**能力**（capability），不映射任何命令语法。OfficeCLI（参考源码 `/tmp/office_research/OfficeCLI`，`SKILL.md` + `src/officecli/Handlers/{Word,Excel,Pptx}` + `schemas/help/`）的能力清单是长期对齐的北极星；但 AIOffice 是 100% 自研（C#/.NET，DocumentFormat.OpenXml + ClosedXML），命令与参数**刻意不**与上游 1:1 还原 —— 我们的对外表面是自己的 14 动词 AI 原生命令面（见 `docs/DESIGN.md`；M1 起含 `preview`）。
>
> 账本与代码之间的契约：任何标记为未实现的能力，运行时必须返回类型化 `unsupported_feature` envelope（`suggestion` 必须给出 workaround），**绝不 crash、绝不静默 no-op**。

## 图例

| 标记 | 含义 |
|------|------|
| ✅ M0 已实现 | M0 脚手架交付：13 动词命令面 + MCP，基础元素读写 |
| ✅ M1 已实现 | M1（v0.2.0）交付：页眉页脚、xlsx 图表（bar/line/pie）、pptx 母版/版式只读、PNG 渲染、实时预览 + 选区回读 |
| 🔜 M1 | M1 计划但尚未落地的剩余项（顺延至 M2 窗口） |
| 📋 M2 | 对象层扩展：图表、批注、条件格式、域/TOC、透视表前置等 |
| 🗺 M3 | 深水区：修订、动画、透视表、OLE/3D/SmartArt、RTL、dump 回放 |
| ❌ 不计划 | 与 AIOffice 设计冲突，或属上游生态/语法专有 |

**M0 元素覆盖范围**（诚实声明）：docx = paragraph / run / table / tr / tc（文本 + 基础字符格式 + 内置样式套用）；xlsx = sheet / cell / range（值、公式、基础字符格式）；pptx = slide / shape / textbox（文本、位置、尺寸）+ 标题占位符。其余元素类型在 M0 一律返回 `unsupported_feature`。

---

## 0. 跨格式通用能力

| 能力 | 说明 | AIOffice 状态 | 实现备注 |
|---|---|---|---|
| 新建空白文档 | docx / xlsx / pptx，按扩展名推断类型，可设标题 | ✅ M0 已实现 | `WordprocessingDocument.Create` / ClosedXML `XLWorkbook` / `PresentationDocument.Create` |
| 结构化读取视图 | outline / text / stats / structure 四种视图，支持范围与字节上限 | ✅ M0 已实现 | 各格式 handler 遍历 OpenXml DOM / ClosedXML 对象模型 |
| 文档体检（issues/lint） | 格式/内容/结构问题清单，附修复建议 | ✅ M0 已实现 | `validate` 内置自研 lint；M0 基础规则集，规则库随能力扩展 |
| annotated 注记文本视图 | 纯文本 + 行内格式注记 | 🔜 M1 | 在 text 视图管线上叠加格式标注 |
| query 基础选择器 | `[attr=value]`、`:contains("x")`、路径前缀限定；返回稳定路径 | ✅ M0 已实现 | 自研选择器解析 + 各格式节点枚举 |
| query 高级选择器 | `!=` / `~=` / `>=` / `<=`、`:empty` / `:has()` / `:no-alt`、布尔 `and`/`or`（含括号）、子代组合器 `>` | 🔜 M1 | 选择器 AST 扩展，M0 解析失败时报 `invalid_args` + 受支持语法清单 |
| 稳定 ID 寻址 | `@paraId` / `@id` / `@name` 形式路径，插删后不漂移 | 🔜 M1 | docx w14:paraId、pptx shape id/name；M0 仅 1-based 位置索引 |
| get 单节点 + 属性 | 读取一个节点及全部有效属性 | ✅ M0 已实现 | 各格式 NodeBuilder 输出统一节点 JSON |
| 原子批量编辑 | 多 op 一次校验一次落盘，全成或全败，`--dry-run` 预演 | ✅ M0 已实现 | `edit --ops`：内存中应用全部 op 后一次 Save |
| set / add / remove / move 四原语 | 路径定位的属性修改、元素插入（after/before/index）、删除、移动 | ✅ M0 已实现 | 仅覆盖 M0 元素类型，其余 `unsupported_feature` |
| swap 元素互换 | 两元素互换位置 | 📋 M2 | 可由两次 move 组合表达，专用 op 后置 |
| 元素克隆 | 从既有元素复制（含图片/图表等跨 part 关系） | 📋 M2 | OpenXml `CloneNode(true)` + 跨 part 关系重建 |
| 文本锚点定位插入 | 按文本匹配定位插入点；行内类型段内插入，块类型自动劈段 | 📋 M2 | 自研 run 切分（依赖跨 run 文本匹配） |
| 全文 find/replace | 字面量/正则，匹配跨 run 边界，作用域由路径控制 | 🔜 M1 | 自研跨 run 文本映射表；返回 `matched` 计数 |
| find 命中文本格式化 | 只对匹配文本应用格式（自动劈 run） | 📋 M2 | 依赖 find/replace 的 run 切分基建；上游 xlsx 也不支持 |
| 模板合并 | `{{key}}` 占位符 + JSON 数据，三种格式的文本 run 内替换 | ✅ M0 已实现 | `template` 动词；跨 run 占位符拼接 |
| CSV/TSV 导入 xlsx | 表头识别（AutoFilter + 冻结首行）、起始单元格 | 📋 M2 | ClosedXML `InsertTable` / 逐格写入 |
| dump 可回放导出 | 整文档导出为可回放的 edit-ops JSON（round-trip） | 🗺 M3 | 各格式 BatchEmitter 等价物；上游 docx 全覆盖是高水位线 |
| render HTML | docx/xlsx 整文档 HTML；pptx 每页 SVG-or-HTML | ✅ M0 已实现 | 自研渲染器（上游 HtmlPreview 系列为行为参考） |
| render PNG 截图 | 整页/网格缩略图 PNG | ✅ M1 已实现 | `AIOffice.Render`：html/svg → headless Chrome/Edge 截图；docx/xlsx 整文档、pptx 单页（`--scope /slide[N]`，缺省首页 + warning）；网格 `--grid` 后置 |
| render PDF | 导出 PDF | 🗺 M3 | 上游经插件实现；评估 html→pdf 通道 |
| OpenXML schema 校验 | 对照 OOXML schema 报错误清单 | ✅ M0 已实现 | `OpenXmlValidator`（也是所有 mutating 测试的 oracle） |
| 实时预览（watch） | 本地 server，文件变更自动刷新浏览器 | ✅ M1 已实现 | `AIOffice.Preview`：`preview open/close` 本地 server（SSE 自动刷新）；MCP `preview_open` 拉起 detached 子进程 |
| 浏览器选区回读 | 用户在预览中点选元素，CLI/MCP 读取选中路径 | ✅ M1 已实现 | 渲染产物全量 `data-aio-path` 标注（docx 块级 / xlsx 单元格 / pptx 形状）+ `POST /selection`；`preview selection` / MCP `preview_selection` 读回 |
| mark 待审修改提案 | 高亮 + 备注 + tofix 的"待人审"标记管线 | 🗺 M3 | 依赖 watch 进程态；与永久批注（comment）区分 |
| goto 视图定位 | 滚动预览浏览器到指定元素 | 🗺 M3 | watch IPC |
| 常驻进程模式 | 文件常驻内存、显式 open/close/save | ❌ 不计划 | 无状态进程 + rev 守卫 + 快照已覆盖正确性；性能实测有瓶颈再议 |
| raw XML 直改 | 按 part + XPath 直接增删改 XML、新建 part | 🗺 M3 | L2 覆盖不足时的逃生舱；受 sandbox 显式开关门禁 |
| MCP stdio server | 与 CLI 同一命令面的 14 工具（M1 起注册 `preview_open`/`preview_selection`） | ✅ M0 已实现 | ModelContextProtocol 官方 C# SDK；CLI/MCP 共用内部命令层 |
| 渐进式 help | 寻址语法、选择器、逐元素属性文档 | ✅ M0 已实现 | `help [topic]`，内容由 schema 生成 |
| 机器可读 schema | 整个命令面的 JSON 描述，agent 自省代替猜测 | ✅ M0 已实现 | `schema [verb]`；超出上游（其仅有元素级 help --json） |
| doctor 环境体检 | runtime / workspace / 快照目录诊断 | ✅ M0 已实现 | 自研检查项 |
| 快照环 + 恢复 | 每次 mutating 前自动快照（保留 20），可 list/restore | ✅ M0 已实现 | 文件级复制环；上游完全没有撤销机制 |
| rev 并发守卫 | `meta.rev`（SHA256 前 12 hex）；`--expect-rev` 不匹配在写前拒绝 | ✅ M0 已实现 | 防外部并发改写；上游没有 |
| workspace sandbox | 文件参数 realpath + symlink 逃逸检查，越界 `sandbox_denied` | ✅ M0 已实现 | `--workspace` / `AIOFFICE_WORKSPACE`，默认 cwd |
| 文档保护 / 密码 | docx protection=forms、xlsx workbook password | 📋 M2 | OpenXml `DocumentProtection` / ClosedXML `Protect` |
| 插件协议 | 外部格式扩展（.doc / .hwpx / pdf-exporter） | ❌ 不计划 | 上游生态专有；AIOffice 聚焦 OOXML 三格式 |
| skills 分发 / load_skill | 向 agent 安装 SKILL.md、按场景加载技能 | ❌ 不计划 | 属 agent 侧生态；AIOffice 用 help/schema 自描述替代 |

---

## 1. docx 能力

| 能力 | 说明 | AIOffice 状态 | 实现备注 |
|---|---|---|---|
| 段落 | 增删改移、文本、对齐/缩进/间距基础属性 | ✅ M0 已实现 | `Wordprocessing.Paragraph` / `ParagraphProperties` |
| 字符格式基础 | 粗体/斜体/下划线/字号/颜色/字体 | ✅ M0 已实现 | `Run` / `RunProperties` |
| 套用内置样式 | style=Heading1/Title/Quote 等引用 | ✅ M0 已实现 | `ParagraphStyleId`；create 时注入最小默认样式库 |
| 样式定义与修改 | 新建/改写 style，行距、缩进、段边框、styles 库管理 | 🔜 M1 | `StyleDefinitionsPart` |
| 多文种字体槽位 | font.latin/ea/cs、lang 槽位、bold.cs/size.cs 等 CS 变体 | 📋 M2 | `RunFonts`、`Languages` |
| 表格基础 | 建表（rows×cols）、加行、单元格文本读写 | ✅ M0 已实现 | `Table` / `TableRow` / `TableCell` |
| 表格格式 | 边框/底纹/列宽/对齐/表样式引用 | 🔜 M1 | `TableProperties` / `TableCellProperties` |
| 单元格合并 | 水平 gridSpan、垂直 vMerge、hMerge | 🔜 M1 | `GridSpan` / `VerticalMerge` |
| 虚拟列操作 | 把 `/tbl/col[N]` 当实体：整列增/删/移/拷 | 📋 M2 | 自研：跨行遍历同列 `TableCell`（上游 table-column schema 为参考） |
| 图片 | 插入、尺寸、alt 文本、二进制导出 | 🔜 M1 | `ImagePart` + `Drawing`（DrawingML inline/anchor） |
| 文本框 / 形状 | 几何、填充、线条、环绕、锚定 | 📋 M2 | DrawingML `wps:wsp` |
| 页眉 / 页脚 | 增删改，首页/奇偶页变体；寻址 `/header[1]/p[1]` | ✅ M1 已实现 | `HeaderPart` / `FooterPart` + `SectionProperties` 引用；read structure / render 同步暴露 |
| 节属性 | 纸张/边距/方向/分栏/页码格式枚举/页面边框 | 🔜 M1 | `SectionProperties`（`PageSize`/`PageMargin`/`Columns`/`PageNumberType`） |
| 分页符 / 换行符 | page break、line break、column break | 🔜 M1 | `Break`（`BreakValues`） |
| 列表与多级编号 | 项目符号、编号、多级列表（num/abstractNum/lvl） | 🔜 M1 | `NumberingDefinitionsPart` |
| 超链接 | 外链 + 锚点链接 | 🔜 M1 | `Hyperlink` + `HyperlinkRelationship` |
| 书签 | bookmarkStart/End，配合交叉引用 | 📋 M2 | `BookmarkStart` / `BookmarkEnd` |
| 批注 | 永久批注（区别于 mark 临时提案），读写删 | 📋 M2 | `WordprocessingCommentsPart` |
| 脚注 / 尾注 | 增删改 + `get /` 暴露清单 | 📋 M2 | `FootnotesPart` / `EndnotesPart` |
| 域（field） | PAGE/NUMPAGES/REF/SEQ/DATE 等（上游覆盖 28 种）；fieldchar/instrtext 低层 | 📋 M2 | `SimpleField` / `FieldChar` + `FieldCode`；常用域优先 |
| 目录 TOC | 插入 TOC 域 + 样式骨架 | 📋 M2 | TOC field + `SdtBlock` 容器 |
| 派生字段重算 | TOC 页码 / PAGE / NUMPAGES / 交叉引用刷新 | 🗺 M3 | 无 Word 后端可依赖：headless 渲染估算页码（上游非 Windows 同样是估算） |
| 图表 | chart + chart-axis / chart-series 子实体 | 📋 M2 | `ChartPart`（DrawingML charts，与 xlsx/pptx 共享实现） |
| 公式（equation） | OMML 数学公式读写 | 🗺 M3 | `DocumentFormat.OpenXml.Math` |
| 水印 | 文本/图片水印 | 📋 M2 | 页眉内 VML shape |
| 表单域 | formfield 读写 + 表单字段清单导出 | 🗺 M3 | `FormFieldData` 系列 |
| 内容控件（sdt） | 结构化文档标签读写 | 🗺 M3 | `SdtBlock` / `SdtRun` |
| 修订（track changes） | ins/del/format/moveFrom/moveTo 写入；accept/reject；按作者过滤；tracked find&replace | 🗺 M3 | `InsertedRun` / `DeletedRun` / rPrChange 系列；上游此处行为最全，逐项对照 |
| 编辑权限范围 | permStart/permEnd 区域授权 | 🗺 M3 | `PermStart` / `PermEnd` |
| OLE 对象 | 嵌入对象读写、二进制导出 | 🗺 M3 | `EmbeddedObjectPart` |
| 制表位 | tabs 简写、ptab | 📋 M2 | `Tabs` / `PositionalTab` |
| 文档默认值 | docDefaults 默认字体/字号 | 🔜 M1 | `DocDefaults`（StylesPart 内） |
| 文档设置 | autoHyphenation、evenAndOddHeaders、docGrid、CJK spacing、protection=forms | 📋 M2 | `SettingsPart`（`Settings` 子元素族） |
| RTL 与 locale | direction=rtl（段/run/表/节）、rtlGutter、按 locale 设分文种默认字体与自动 RTL | 🗺 M3 | `BiDi` / `RightToLeftText` + locale→字体映射表 |

---

## 2. xlsx 能力

| 能力 | 说明 | AIOffice 状态 | 实现备注 |
|---|---|---|---|
| 工作表管理 | 增/删/改名/移动；改名后公式引用级联更新 | ✅ M0 已实现 | ClosedXML `IXLWorksheet`（rename 自动级联引用） |
| 工作表属性 | visible/hidden/veryHidden、sheetView 基础 | 🔜 M1 | `IXLWorksheet.Visibility` |
| 打印设置 | 边距、printTitleRows/Cols、手动分页符（row/colbreak） | 📋 M2 | `IXLPageSetup` / `PageBreaks` |
| 单元格读写 | 值（类型推断：数字/日期/布尔/文本）、公式写入（`=` 自动识别） | ✅ M0 已实现 | `IXLCell.Value` / `FormulaA1` |
| 区域寻址与批量赋值 | `/Sheet1/A1:C10`、`/Sheet1/row[3]`、带空格表名 `/'Q3 Data'/B2` | ✅ M0 已实现 | `IXLRange`；自研寻址解析 |
| 字符格式基础 | 粗/斜/字号/字色/填充色 | ✅ M0 已实现 | `IXLStyle.Font` / `Fill` |
| 数字格式 / 对齐 / 边框 | 内置与自定义 number format、对齐、边框样式 | 🔜 M1 | `IXLStyle.NumberFormat` / `Alignment` / `Border` |
| 富文本单元格 | 单元格内多 run 异格式 | 🔜 M1 | `IXLCell.GetRichText()` |
| 行列操作 | 插入/删除/隐藏、行高列宽；插删后公式引用自动重写 | 🔜 M1 | `InsertRowsAbove` / `Delete` 等（ClosedXML 自动调整引用） |
| 单元格 shift 语义 | 删除补位（left/up）、插入挤位（right/down），对齐 Excel UI 对话框 | 📋 M2 | `InsertCellsAfter` / `Delete(XLShiftDeletedCells)` |
| 合并单元格 | merge/拆分；query 枚举 merge/mergedrange | 🔜 M1 | `IXLRange.Merge()` / `Worksheet.MergedRanges` |
| 公式求值（回读） | 读取公式单元格时返回计算值（常用函数集） | ✅ M0 已实现 | ClosedXML CalcEngine；不支持的函数→`formula_not_evaluated` warning + 建议（绝不静默给空值） |
| 现代函数写入前缀 | XLOOKUP/FILTER/LET/LAMBDA 等写入时自动加 `_xlfn.`（FILTER 加 `_xlfn._xlws.`）、spill 引用 `A1#` | 📋 M2 | 自研 qualifier（上游 `ModernFunctionQualifier` 为行为参考）；保证 Excel 打开能算 |
| 动态数组 / 迭代求解求值 | FILTER/SORT/UNIQUE/SEQUENCE 求值；RATE/IRR 迭代解 | 🗺 M3 | 上游同样不求值（写入可用）；评估自研迭代器的性价比 |
| 排序 | 多列 `COL DIR` 列表、表头感知；拒绝含合并/公式区域；超链接/批注/条件格式随行迁移 | 📋 M2 | `IXLRange.Sort`；旁挂元数据跟随需自研 |
| 自动筛选 | 区域 AutoFilter | 🔜 M1 | `Range.SetAutoFilter()` |
| 冻结窗格 | 冻结行/列 | 🔜 M1 | `SheetView.Freeze(rows, cols)` |
| 命名区域 | 定义/引用/删除；`[@name=X]` 寻址 | 🔜 M1 | `IXLWorkbook.NamedRanges` |
| 表（listobject） | 区域转表、内建表样式目录、totals row | 📋 M2 | `IXLTable` + `XLTableTheme` |
| 条件格式 | databar / colorscale / iconset / cellIs / formulacf / topN / aboveAverage / containstext / dateoccurring / duplicate / unique（上游 11 类全） | 📋 M2 | `IXLConditionalFormat`（ClosedXML 覆盖大部分类型） |
| 数据验证 | 列表/范围/自定义公式验证 + 提示语 | 📋 M2 | `IXLDataValidation` |
| 迷你图 | line / column / winloss sparkline 组 | 📋 M2 | `IXLSparklineGroups` |
| 图表 | bar/line/pie 增删读（`/Sheet1/chart[1]` 寻址、anchor 简写）；轴/系列子实体逐属性、pareto 后置 | ✅ M1 已实现 | **ClosedXML 不支持图表**：自研 OpenXml `ChartPart` + `DrawingsPart`（`ExcelCharts.cs`）；轴/系列深编辑 → 📋 M2 |
| 图片 | 插入/定位/缩放（PNG/JPEG）；SVG | 🔜 M1 | `IXLPicture`；SVG 需 OpenXml 双 part（svgBlip + png fallback）→ 📋 M2 |
| 批注 | 单元格批注读写删 | 📋 M2 | `IXLCell.CreateComment` |
| 形状 / 文本框 | 浮动 shape/textbox；选择器需枚举 grpSp 内叶子 | 🗺 M3 | ClosedXML 不支持：OpenXml `DrawingsPart`（xdr:sp） |
| 切片器 | 表/透视表切片器 | 🗺 M3 | OpenXml `SlicerPart`（ClosedXML 无） |
| 数据透视表 | rows/cols/values（Field:func[:showDataAs]）、filters、layout、grandTotals/subtotals、topN、calculatedField、日期自动分组、缓存共享 | 🗺 M3 | `IXLPivotTable` 起步（覆盖布局与聚合 11 种）；calculatedField/缓存 CoW 需 OpenXml 补 |
| 超链接 | 单元格超链接读写 | 🔜 M1 | `IXLCell.SetHyperlink` |
| 工作簿级属性 | password、calc.mode=manual、calc.refMode=r1c1 | 📋 M2 | `Protect` / `CalculateMode` / OpenXml `WorkbookProperties` |
| CSV/TSV 导入 | 文件/stdin、表头→AutoFilter+冻结、起始单元格 | 📋 M2 | 自研解析 + ClosedXML 写入（见跨格式"CSV 导入"行） |
| 行按列名查询 | `row[Salary>5000]` 式按表头列名过滤 | 📋 M2 | 自研：表头行映射 + 选择器扩展 |
| OLE 对象 | 嵌入对象、preview 图 | 🗺 M3 | OpenXml `EmbeddedObjectPart` |
| RTL | sheetView rightToLeft、cell/comment direction | 🗺 M3 | OpenXml `SheetView.RightToLeft` |
| 拼音指引（phonetic） | 东亚文本 phonetic runs | 🗺 M3 | OpenXml `PhoneticRun` |

---

## 3. pptx 能力

| 能力 | 说明 | AIOffice 状态 | 实现备注 |
|---|---|---|---|
| 幻灯片增删改移 | 新建（基于默认版式）、删除、重排 | ✅ M0 已实现 | `SlidePart` + `SlideIdList` 维护 |
| 隐藏幻灯片 | show=false | 🔜 M1 | `Slide.Show` |
| 幻灯片克隆 | 复制 slide 含全部跨 part 关系（图片/图表/媒体） | 📋 M2 | `AddPart` + 关系 id 重映射 |
| 形状 / 文本框 | 新建、位置/尺寸（EMU 及 cm/pt/px 单位换算）、文本读写 | ✅ M0 已实现 | `P.Shape` + `a:txBody`；`Transform2D` |
| 文本段落与 run 格式 | 字体/字号/颜色/粗斜/对齐/项目符号基础 | ✅ M0 已实现 | `a:p` / `a:r` / `a:rPr`（M0 覆盖基础属性） |
| 标题占位符填充 | 新片设标题（create --title、slide title 属性） | ✅ M0 已实现 | 版式 placeholder 匹配（phType=title/ctrTitle） |
| 任意占位符 | phType=body/subTitle/footer/dt/sldNum 等定位与填充 | 🔜 M1 | `PlaceholderShape` 索引/类型匹配 |
| 形状外观基础 | 纯色填充、线条颜色/粗细、预设几何（rect/ellipse/arrow 等） | 🔜 M1 | `a:solidFill` / `a:ln` / `a:prstGeom` |
| 形状外观高级 | 渐变/图片填充、阴影/发光效果、旋转、图片亮度对比度 | 📋 M2 | `a:gradFill` / `a:effectLst` / `xfrm@rot` |
| effective 继承值溯源 | 读出属性的最终生效值及来源（shape→layout→master 链） | 🗺 M3 | 自研继承解析器（上游 effective.X.src 为行为参考） |
| 母版 / 版式寻址 | `/master[1]`、`/master[1]/layout[2]` 路径，get/query 只读 | ✅ M1 已实现 | `SlideMasterPart` / `SlideLayoutPart` 枚举；编辑仍 `unsupported_feature`（→ M2） |
| 母版 / 版式编辑 | 对 master/layout 上元素 add/set/remove | 📋 M2 | 同 slide 编辑管线复用 |
| 图片 | 插入/位置/尺寸/alt；SVG、裁剪、滤镜属性后置 | 🔜 M1 | `ImagePart` + `P.Picture`；SVG/滤镜 → 📋 M2 |
| 表格 | graphicFrame 建表、行列增删、单元格文本/合并、内建样式目录、列实体操作 | 📋 M2 | `a:tbl`（DrawingML table） |
| 图表基础 | 柱/线/饼 + 轴/系列子实体 | 📋 M2 | `ChartPart`（与 docx/xlsx 共享 DrawingML chart 层） |
| 图表高级 | pieOfPie/barOfPie、轴线/网格线逐属性、dropLines/hiLowLines/upDownBars、图表动画（chartBuild） | 🗺 M3 | ChartPart 深水区 |
| 连接线 | from/to 锚接形状（接受 @name）、起止端点改绑 | 📋 M2 | `P.ConnectionShape` + `a:stCxn`/`a:endCxn` |
| 组合（group） | 建组/解组、get/query 深度遍历、link/tooltip | 📋 M2 | `P.GroupShape` 递归 |
| 备注（notes） | 读写每页演讲者备注 | 🔜 M1 | `NotesSlidePart` |
| 批注（legacy） | 经典批注读写删 | 📋 M2 | `SlideCommentsPart` |
| 现代线程批注 | p188 modern comments（线程/回复）round-trip | 🗺 M3 | `powerPointComments` part（OpenXml 3.x 已建模） |
| 动画 | 进入/强调/退出预设（上游 15 emphasis + 16 exit）、多效果链、motion path、repeat/restart/autoReverse | 🗺 M3 | `p:timing` 树手写（OpenXml 无高层 API） |
| 切换（transition） | 常规切换预设（fade/push/wipe + p15 12 种） | 📋 M2 | `P.Transition` + p14/p15 扩展命名空间 |
| Morph 平滑切换 | p14 morph + `!!` 命名匹配感知 | 🗺 M3 | p14 扩展 + 对象名约定 |
| 媒体（video/audio） | 嵌入/链接、loop、autoStart | 🗺 M3 | `VideoReferenceRelationship` + `p:nvPr` media |
| 公式（equation） | OMML 公式形状 | 🗺 M3 | `a14:m` 内嵌 Math |
| zoom | summary / section / slide zoom 跳转对象 | 🗺 M3 | p15 `sectionZoom`/`slideZoom` |
| OLE 对象 | 嵌入对象 + preview 图 | 🗺 M3 | `graphicFrame` + `p:oleObj` |
| 3D 模型 | glTF 模型嵌入、rotation=ax,ay,az | 🗺 M3 | am3d 扩展 part |
| SmartArt | 图示读写（data/layout/colors part 组） | 🗺 M3 | `DiagramDataPart` 等四件套 |
| 超链接 | 文本/形状级跳转（外链、页内跳转）、tooltip | 🔜 M1 | `a:hlinkClick` + 关系 |
| 主题 / 文档级属性 | theme 颜色字体、defaultFont、show.loop、print.what | 📋 M2 | `ThemePart` / `PresentationPropertiesPart` |
| RTL | shape/notes direction、图表轴 RTL | 🗺 M3 | `a:bodyPr@rtlCol` / `a:pPr@rtl` |
| 逐页渲染 SVG/HTML | 每页 slide 输出 SVG（或 HTML 退化） | ✅ M0 已实现 | 自研渲染器（基础形状/文本；覆盖面随元素能力增长） |
| 截图网格 / 页范围 | PNG 缩略图网格（--grid）、html 页范围 | 🔜 M1 | 单页 PNG 已落地（`render --to png --scope /slide[N]`）；网格拼接与页范围未实现 |

---

## 4. 统计

| 范围 | ✅ 已实现（M0+M1） | 🔜 M1 | 📋 M2 | 🗺 M3 | ❌ 不计划 | 合计 |
|---|---|---|---|---|---|---|
| 跨格式通用 | 20 | 4 | 6 | 5 | 3 | 38 |
| docx | 5 | 9 | 12 | 8 | 0 | 34 |
| xlsx | 6 | 10 | 12 | 7 | 0 | 35 |
| pptx | 6 | 7 | 10 | 12 | 0 | 35 |
| **合计** | **37** | **30** | **40** | **32** | **3** | **142** |

> 解读：M0 的 31 项里有相当一部分是 AI 原生层（envelope、快照、rev、sandbox、schema、MCP）—— 这是 AIOffice 的差异化地基，上游完全没有。格式能力本身（docx 4 / xlsx 5 / pptx 5）在 M0 刻意收窄到"段落、单元格、形状"三个最小核心，换取每一项都带测试、过 `OpenXmlValidator`、守住 round-trip 字节一致律。
>
> M1（v0.2.0）新增 6 项：docx 页眉/页脚、xlsx 图表（bar/line/pie，自研 ChartPart）、pptx 母版/版式只读寻址、跨格式 PNG 渲染（headless 浏览器）、实时预览 server、浏览器选区回读（`data-aio-path` 契约，原排 M2 提前落地）。其余 102 项按 M1→M3 排队，每一项落地时同步更新本表状态与"实现备注"中的真实 API。
