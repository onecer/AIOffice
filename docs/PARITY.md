# AIOffice 能力账本（Capability Ledger）

> 本账本只记录**能力**（capability），不映射任何命令语法。OfficeCLI（参考源码 `/tmp/office_research/OfficeCLI`，`SKILL.md` + `src/officecli/Handlers/{Word,Excel,Pptx}` + `schemas/help/`）的能力清单是长期对齐的北极星；但 AIOffice 是 100% 自研（C#/.NET，DocumentFormat.OpenXml + ClosedXML），命令与参数**刻意不**与上游 1:1 还原 —— 我们的对外表面是自己的 16 动词 AI 原生命令面（见 `docs/DESIGN.md`；M1 起含 `preview`，M8 起含 `diff`）。
>
> 账本与代码之间的契约：任何标记为未实现的能力，运行时必须返回类型化 `unsupported_feature` envelope（`suggestion` 必须给出 workaround），**绝不 crash、绝不静默 no-op**。

## 图例

| 标记 | 含义 |
|------|------|
| ✅ M0 已实现 | M0 脚手架交付：13 动词命令面 + MCP，基础元素读写 |
| ✅ M1 已实现 | M1（v0.2.0）交付：页眉页脚、xlsx 图表（bar/line/pie）、pptx 母版/版式只读、PNG 渲染、实时预览 + 选区回读 |
| ✅ M2 已实现 | M2（v0.3.0）交付：docx 修订（文本级）/批注/样式管理/图片、xlsx 透视表/条件格式/图片、pptx 背景/备注/图片、大文件守卫 |
| ✅ M3 已实现 | M3（v0.4.0，功能第一）交付：docx 列表/链接/书签/脚注/页面设置/格式修订处理/批注回复、xlsx 流式读取/散点+面积图/命名区域/冻结/筛选/打印设置、pptx 图表/切换/几何/z 序、跨文档 dataFrom、PDF 渲染、大小上限 opt-in |
| ✅ M4 已实现 | M4（v0.5.0）交付：三格式共享 find/replace 契约（`op:replace`，零命中 = `find_no_match` warning；文档级 `"/"` 作用域展开 + CLI `--find/--replace` 糖 + 聚合 `{replacements, locations}`，docx `--track` 生成 del+ins 修订对）；docx TOC/水印/尾注/插入分节符；xlsx 批量 2D 写入（锚点/区域两形态 + >50k 空白 sheet SAX 流式写）、行列插删/行高列宽/隐藏（`col[C]` 字母寻址）、单元格批注；pptx 进入动画（appear/fade/flyIn/wipe）、经典批注、嵌入式图表工作簿（PowerPoint「编辑数据」可用 + `embedData:true` 退化改造）；tag 触发的 GitHub release 自动化 |
| ✅ M5 已实现 | M5（v0.6.0，markdown/csv 桥）交付：`create --from`（.md → .docx Markdig 全结构导入 / .csv → .xlsx 类型化导入）+ `read --view markdown`（docx → GFM 导出，与导入结构 round-trip）/ `--view csv`（单 sheet RFC 4180 导出）；docx 深表格（合并/边框/底纹/列宽/headerRow）、域（PAGE/NUMPAGES/DATE/TITLE + leadingText）、首页/奇偶页眉页脚变体；xlsx 数据验证（list 下拉 + wholeNumber/decimal/date/textLength）、迷你图（line/column/winLoss）、线程批注（xl/threadedComments + 回复）、单元格超链接（外链/`#Sheet!A1` 内链）；pptx 原生表格（合并 + light/medium/dark 样式）、强调/退出动画（pulse/grow/spin/colorPulse + fadeOut/flyOut/wipeOut）、批注回复线程（p15 threadingInfo）、SmartArt 只读 |
| ✅ M6 已实现 | M6（v0.7.0，深水区攻坚）交付：docx **公式**（LaTeX → OMML 自研转换器，行内 `$…$` / 显示 `$$…$$`、矩阵/分式/根式/上下标/希腊字母/大型算符，未知命令降级为 `equation_partial` warning，原始 LaTeX 存档可忠实回读）、**右到左/双向**（段落 `w:bidi` / run `w:rtl` / 表格 `w:bidiVisual`）、**多栏分节**（columns/columnGap + columnBreak）；xlsx **就地流式写入**（`stream:true` 或 >20 MB 文件，整批可流式时经 SAX writer 原地改写大工作簿，深单元格/批量写秒级完成）、**Excel 表（ListObject）**（命名 + 内置样式 + 汇总行 + 结构化引用 `=SUM(Sales[Amount])` 求值）、**大纲分组**（`row[a]:row[b]` / `col[a]:col[b]` 行列折叠分组）；pptx **母版/版式编辑**（背景 + 主题强调色 + 克隆版式 + 母版/版式形状）、**幻灯片分节**（p14:sectionLst，按 sldId 跟踪）、**幻灯片尺寸/宽高比**（16:9/4:3/自定义）、**动画时间轴重排**（move/set 重排重调）；新增寻址形式 `/`（演示文稿根）、`/body/p[i]/omath[j]`、`/section[i]`、`row[a]:row[b]` 跨段、可编辑 `/master[1]/layout[i]` |
| ✅ M7 已实现 | M7（v0.8.0，交付前审计）交付：三格式共享 **`audit` 动词 + `office_audit` MCP 工具（第 15 个工具）**—— 无障碍 + 质量 lint，findings 是**数据**（`ok:true`/exit 0，即使有 error 级 finding），`--category accessibility\|quality\|all` + `--severity error\|warning\|info`（最低上报级）+ `--fix`（仅安全自动修复：占位 alt 文本、标记表头行、设置文档/幻灯片标题、删除孤儿书签，修复后重审返回 `{fixed:N, remaining:[…]}`）；codes：`a11y_no_alt_text`/`a11y_no_table_header`/`a11y_no_doc_title`/`a11y_no_slide_title`/`a11y_heading_skip`/`a11y_low_contrast`/`a11y_tiny_font`/`a11y_merged_data_cells`/`a11y_reading_order` + `quality_broken_ref`/`quality_formula_error`/`quality_broken_link`/`quality_empty_heading`/`quality_off_canvas`/`quality_empty_placeholder`/`quality_orphan_bookmark`/`quality_duplicate_id`。docx **文档属性**（`/properties` set + `read --view properties`：core title/subject/author… + 类型化 custom string/number/bool/date）、**内容控件**（`add type:contentControl` text/dropdown/date/checkbox，`/sdt[@tag=X]` set/get、`read --view fields`）；xlsx **命名单元格样式**（`add type:cellStyle` 定义 numberFormat/bold/fill/color/border，`set {cellStyle:"X"}` 套用到区域，`read --view styles` / `/style[@name=X]`）、**文档属性**（`/properties`）；pptx **alt-text/标题/阅读顺序**审计 + 显式 `altText`/`altTitle` shape prop、**文档属性**（`/properties`）；新增 `office_read` view 枚举 `properties`/`fields`/`styles` |
| ✅ M8 已实现 | M8（v0.9.0，对比与审阅）交付：三格式共享 **`diff` 动词 + `office_diff` MCP 工具（第 16 个工具）**—— 语义对比当前文档与基线（另一同格式文件 `other` 或自身快照 `--snapshot N`，跨格式 → `invalid_args`），返回**已排序、跨平台确定性**的变更清单 `{changes:[{kind:added\|removed\|modified\|moved, path, before?, after?, detail?}], summary:{added,removed,modified,moved}, baseline}`（按 `(path, kind)` 序贯字典序排序；LCS/内容哈希区分 moved 与 added+removed），变更是**数据**（`ok:true`/exit 0），`--view summary` 裁剪为 path+kind、`detailed`（默认）保留完整 before/after；快照 diff 工作流（编辑自动快照 → `diff f --snapshot 1` = 上次编辑的变更集）。docx **题注/交叉引用**（`add type:caption` Figure/Table/Equation + SEQ 字段 + 书签，寻址 `/caption[@label=Figure][i]`；`add type:crossRef` REF 字段 + `labelAndNumber\|numberOnly\|text\|page`，编号缓存 `caption_numbers_cached` warning）；xlsx **切片器**（`add type:slicer` 表列 / 透视字段，原始 OpenXml 自研，寻址 `/Sheet1/slicer[i]` / `slicer[@name=X]`）；pptx **形状超链接/动作**（`set {hyperlink}` 外链 url / 跳转 `#slide:N` / 放映动作 `#first\|#last\|#next\|#prev\|#end`、`set {linkText}` 文本超链接，`get` 回读规范形式）；新增寻址 `/caption[@label=…][i]`、`/Sheet1/slicer[i]` |
| 🔜 M1 | M1 计划但尚未落地的剩余项（顺延至后续窗口） |
| 📋 M2 | M2 计划但尚未落地的剩余项（顺延至后续窗口） |
| ✅ M9 已实现 | M9（v0.10.0，pre-1.0 capstone）交付：三格式共享 **`convert` 动词 + `office_convert` MCP 工具（第 17 个工具，MCP 表面从 16 → 17 工具）**—— 跨格式内容互转，按 (源扩展名, 目标扩展名) 路由：同族文本桥 docx↔md / xlsx↔csv（复用既有桥），跨格式经**内容中立模型**（Core 新增 `NeutralDoc`/`NeutralBlock`/`NeutralRun`/`ImportResult` records + `INeutralConvertible` 接口，三 handler 各实现 ExportNeutral/ImportNeutral）docx↔pptx / docx↔xlsx / pptx↔xlsx / 任意↔md（命令层 `NeutralMarkdown` 序列化器），any→pdf/png/svg/html/txt 经 render 层；转换跨格式本质有损，丢失项（动画/图表/精确样式/公式取值/markdown 不支持的颜色）汇入单条 `convert_lossy` warning + `data.dropped`，目标新建（覆盖既有 → `convert_overwrite` warning），同扩展名 → `invalid_args`（「use edit」），未知目标 → `unsupported_feature`（列出受支持目标）；doctor/`office_status` 新增 **capabilities 自省块**（verbs/MCP 工具数、格式、convert 源/目标、render 目标、audit 类别） |
| 🗺 1.0 之后 | 深水区余项：pptx/xlsx 公式（复用 `AIOffice.Word.Equations` 转换器，跨格式经共享转换器）、OLE 对象、插件机制、API 稳定性承诺、现代 xlsx 批注打磨、动画预设++、多效果动画链/motion path、dump 回放 |
| ❌ 不计划 | 与 AIOffice 设计冲突，或属上游生态/语法专有 |

**M0 元素覆盖范围**（诚实声明）：docx = paragraph / run / table / tr / tc（文本 + 基础字符格式 + 内置样式套用）；xlsx = sheet / cell / range（值、公式、基础字符格式）；pptx = slide / shape / textbox（文本、位置、尺寸）+ 标题占位符。其余元素类型在 M0 一律返回 `unsupported_feature`。

---

## 0. 跨格式通用能力

| 能力 | 说明 | AIOffice 状态 | 实现备注 |
|---|---|---|---|
| 新建空白文档 | docx / xlsx / pptx，按扩展名推断类型，可设标题 | ✅ M0 已实现 | `WordprocessingDocument.Create` / ClosedXML `XLWorkbook` / `PresentationDocument.Create` |
| 结构化读取视图 | outline / text / stats / structure 四种视图，支持范围与字节上限 | ✅ M0 已实现 | 各格式 handler 遍历 OpenXml DOM / ClosedXML 对象模型 |
| 文档体检（issues/lint） | 格式/内容/结构问题清单，附修复建议 | ✅ M0 已实现 | `validate` 内置自研 lint；M0 基础规则集，规则库随能力扩展 |
| 无障碍 + 质量审计（audit） | `audit <file> [--category accessibility\|quality\|all] [--severity error\|warning\|info] [--fix]`（MCP `office_audit`，第 15 个工具）：findings 是**数据**（`ok:true`/exit 0，即使有 error 级 finding，与 validate 同），每 finding 带稳定 `id`=`code#path` + severity + category + message + suggestion + autofixable；`--fix` 只应用**安全**自动修复（占位 alt 文本 `(describe this image)`、标记表头行、设文档/幻灯片标题、删孤儿书签），修复后重审返回 `{fixed:N, remaining:[…]}`，绝不破坏内容、修复前自动快照可撤销 | ✅ M7 已实现 | Core 新增 `IAuditor` 接口 + `AuditFinding`/`AuditResult`/`AuditSummary`/`AuditOptions` records；三 handler 各实现（WordHandler.Audit.cs / ExcelHandler.Audit.cs + ExcelAudit.cs / PptxAudit.cs），命令层共享 `AuditVerb`（CLI `audit` 与 MCP `office_audit` 同产物）；codes 见 `aioffice help audit`；语义 diff 已于 M8 交付（见下行），自定义规则插件 → 🗺 M9 |
| 语义对比 / 评审（diff） | `diff <file> [<other>] [--snapshot N] [--view summary\|detailed]`（MCP `office_diff`，第 16 个工具）：语义对比**当前**文档与基线 —— 另一**同格式**文件（`other`，跨格式 → `invalid_args`）或自身**快照**（`--snapshot N`，缺失索引 → `invalid_args` + 可用编号 candidates）；返回**已排序、跨平台确定性**变更清单 `{changes:[{kind:added\|removed\|modified\|moved, path（current 路径；removed 取 baseline 路径）, before?, after?, detail?}], summary:{added,removed,modified,moved}, baseline}`（按 `(path, kind)` 序贯字典序排序，每次每平台一致；段落/幻灯片/行用 LCS/内容哈希区分 moved 与 added+removed）；变更是**数据**（`ok:true`/exit 0）；`--view summary` 裁剪每条为 `{kind, path}`、`detailed`（默认）保留完整 before/after；快照 diff 工作流：每次成功 `edit` 自动快照前像，故 `diff f --snapshot 1` = 上次编辑的变更集 | ✅ M8 已实现 | Core 新增 `IDiffer` 接口 + `DiffChange`/`DiffResult`/`DiffSummary` records（`DiffResult.FromChanges` 序贯排序）；三 handler 各实现（WordHandler.Diff.cs / ExcelHandler.Diff.cs + ExcelDiff.cs / PptxDiff.cs），命令层共享 `DiffVerb`（CLI `diff` 与 MCP `office_diff` 同产物，快照基线还原至临时文件）；跨格式 docx↔pptx 转换已于 M9 交付（见下行） |
| 跨格式互转（convert） | `convert <src> <dest>`（MCP `office_convert`，第 17 个工具）：按 (源扩展名, 目标扩展名) 路由 —— 同族文本桥 **docx↔md**（markdown 桥）/ **xlsx↔csv**（csv 桥）复用既有代码，跨格式经**内容中立模型**（`INeutralConvertible`）：源 handler `ExportNeutral` → 目标 handler `ImportNeutral`，覆盖 **docx↔pptx / docx↔xlsx / pptx↔xlsx / 任意格式↔md**（md 经命令层 `NeutralMarkdown` 序列化器，故任意 office 格式可转 md），**any→pdf/png/svg/html/txt** 经 render 层；目标**新建**（覆盖既有 → `convert_overwrite` warning），同扩展名 → `invalid_args`（「use edit」），未知目标 → `unsupported_feature`（列出受支持目标）；转换跨格式本质有损 —— 丢失项（动画/图表/SmartArt/精确样式/公式取值/markdown 不支持的颜色与下划线）honest 汇入单条 `convert_lossy` warning + `data.dropped`；envelope `data:{from,to,blocksWritten,dropped:[…],written}` | ✅ M9 已实现 | Core 新增 `NeutralDoc`/`NeutralBlock`/`NeutralRun`/`ImportResult` records + `INeutralConvertible` 接口（与 M7 `IAuditor`、M8 `IDiffer` 同模式：Core 定义、三 handler 各实现 ExportNeutral/ImportNeutral）；WordHandler.Neutral.cs / ExcelHandler.Convert.cs / PptxConvert.cs；命令层 `ConvertVerb`（CLI `convert` 与 MCP `office_convert` 同产物，路由 + 双路径沙箱解析）+ `NeutralMarkdown` 序列化器；pptx export 额外 `out dropped` 重载命名 export 侧损失（动画/图表/SmartArt） |
| 文档属性（core + custom） | `set /properties`（core title/subject/author/keywords/category/comments/lastModifiedBy/created/modified/revision + `custom:{…}` 类型化）、`read --view properties` / `get /properties` 回读，custom 值保留 string/number/bool/date 类型 | ✅ M7 已实现（docx/xlsx）；pptx core | docx：`PackageProperties` + `CustomFilePropertiesPart` 类型化 vt:lpwstr/r8/bool/filetime（WordHandler.Properties.cs）；xlsx：ClosedXML `Properties` + custom（ExcelProperties.cs）；pptx：core props（PptxProperties.cs）；`office_read` view 枚举新增 `properties` |
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
| 全文 find/replace | 共享契约 `{op:replace, props:{find, replace, regex?, matchCase?, wholeWord?}}`：字面量/正则（.NET Regex + 2 秒超时），docx/pptx 匹配跨 run 边界（重写保留首个受影响 run 的格式）；作用域 = 任意容器路径，或 `"/"` = 整文档（docx body+全部页眉页脚 / 每个 sheet / 每页 slide 含备注，命令层展开 + 聚合 `{replacements, locations≤20}`）；CLI 糖 `edit f --find X --replace Y [--regex --match-case --whole-word]`；零命中 = ok + `find_no_match` warning | ✅ M4 已实现 | WordHandler.Replace.cs / ExcelHandler.Replace.cs / PptxReplace.cs + 命令层 `ReplaceSugar`（CLI 与 MCP office_edit 共用）；docx `--track` 时逐命中生成 w:del+w:ins 修订对（仅 body 作用域） |
| find 命中文本格式化 | 只对匹配文本应用格式（自动劈 run） | 📋 M2 | 依赖 find/replace 的 run 切分基建；上游 xlsx 也不支持 |
| 模板合并 | `{{key}}` 占位符 + JSON 数据，三种格式的文本 run 内替换 | ✅ M0 已实现 | `template` 动词；跨 run 占位符拼接 |
| 跨文档数据桥（dataFrom） | pptx add chart 接受 `{"dataFrom":"book.xlsx!Sheet1/A1:B5"}`：首列 → 分类、表头行 → 系列名、其余列 → 系列值；区域写错返回 candidates（表名/最近 usedRange） | ✅ M3 已实现 | 命令层展开（`AIOffice.Mcp.CrossDocDataFrom`，CLI edit 与 MCP office_edit 共用）；工作簿必经 workspace 沙箱解析，经 xlsx handler 读取后注入字面量 |
| CSV/TSV 导入 xlsx | `create orders.xlsx --from orders.csv`：RFC 4180（引号/嵌入逗号换行）、分隔符嗅探（`, ; tab \|`）或 `delimiter` 强制、`--title` 命名 sheet；单元格类型化（数字/日期/布尔转换，`007` 类前导零码**保持文本**）；>50k 单元格走 SAX 流式写；M9 起 `convert data.xlsx data.csv` / `convert orders.csv orders.xlsx` 统一入口 | ✅ M5 已实现（M9 convert 统一） | 自研解析（ExcelCsv.cs）+ ExcelHandler.Csv.cs（CreateFrom 钩子，命令层路由 + 沙箱）；M9 `convert` 复用该桥（xlsx↔csv 直转，其它格式经 xlsx）；表头识别（AutoFilter + 冻结首行）→ 📋 余项 |
| markdown 桥（docx） | `create report.docx --from notes.md`：Markdig GFM 导入（标题→样式、粗斜删/行内代码 run、嵌套列表、管道表格、链接/图片/引用/代码块/分隔线，raw HTML 与缺图降级为 warning）；`read --view markdown` 把 body 导出回 GFM，结构与导入 round-trip；M9 起 `convert report.docx report.md` / `convert report.md report.docx` 统一入口，且 **任意 office 格式↔md** 经 `NeutralMarkdown` 序列化器（xlsx/pptx → md 亦可） | ✅ M5 已实现（M9 convert 统一 + 全格式↔md） | Markdig（BSD-2-Clause）AST 走查（WordHandler.Markdown.cs / WordHandler.MarkdownExport.cs）；IFormatHandler `CreateFrom` 默认实现 = `unsupported_feature`；M9 `convert` 的 docx↔md 复用该桥，xlsx/pptx↔md 经内容中立模型 + 命令层 `NeutralMarkdown` |
| dump 可回放导出 | 整文档导出为可回放的 edit-ops JSON（round-trip） | 🗺 M3 | 各格式 BatchEmitter 等价物；上游 docx 全覆盖是高水位线 |
| render HTML | docx/xlsx 整文档 HTML；pptx 每页 SVG-or-HTML | ✅ M0 已实现 | 自研渲染器（上游 HtmlPreview 系列为行为参考） |
| render PNG 截图 | 整页/网格缩略图 PNG | ✅ M1 已实现 | `AIOffice.Render`：html/svg → headless Chrome/Edge 截图；docx/xlsx 整文档、pptx 单页（`--scope /slide[N]`，缺省首页 + warning）；网格 `--grid` 后置 |
| render PDF | 导出分页 PDF；pptx 整副 deck 一页一片 | ✅ M3 已实现 | `AIOffice.Render` PdfRenderer：handler html/svg → headless Chrome `--print-to-pdf --no-pdf-header-footer`；docx/xlsx A4 分页、pptx 按片尺寸 `@page`；无浏览器 → `unsupported_feature` + workaround |
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
| 大文件守卫 | M3 起默认**不限大小**（功能第一指令）；`AIOFFICE_MAX_FILE_MB` 为 opt-in 上限，超限返回 `file_too_large` + 建议；doctor 报告 `limits.maxFileMb`（缺省 `"unlimited"`） | ✅ M2 已实现（M3 翻转默认） | Core `FileSizeGuard`，三 handler 打开链路统一接入；超大 xlsx 读取走 M3 流式路径（见 xlsx 表）；写入流式 → 🗺 M3 |
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
| 样式定义与修改 | 新建/改写/删除自定义 style（bold/color/fontSize/alignment/spacing/basedOn），`read --view styles` 清单 | ✅ M2 已实现 | `StyleDefinitionsPart`（WordHandler.Styles.cs）；`/styles` add + `/style[@id=X]` set/get/remove；行距/段边框深化 → 📋 M2 余项 |
| 多文种字体槽位 | font.latin/ea/cs、lang 槽位、bold.cs/size.cs 等 CS 变体 | 📋 M2 | `RunFonts`、`Languages` |
| 表格基础 | 建表（rows×cols）、加行、单元格文本读写 | ✅ M0 已实现 | `Table` / `TableRow` / `TableCell` |
| 表格格式 | 边框（all/outer/none + 颜色/线宽）、底纹、headerRow（首行加粗 + `w:tblHeader` 跨页重复）、表宽（cm/%）、columnWidths（逐列，`"auto"` 可混用）、对齐、cellPaddingCm | ✅ M5 已实现 | `TableProperties` / `TableCellProperties`（WordHandler.Tables.cs）；表样式引用 → 📋 余项 |
| 单元格合并 | `mergeRight: N`（gridSpan，吃掉右侧单元格）、`mergeDown: N`（w:vMerge restart/continue 链，1 = 拆开）；合并冲突报 invalid_args；被并掉的格从 get/render 消失，HTML 渲染出真 colspan/rowspan | ✅ M5 已实现 | `GridSpan` / `VerticalMerge`（WordHandler.Tables.cs）；hMerge 旧式 → 不计划（gridSpan 等价） |
| 虚拟列操作 | 把 `/tbl/col[N]` 当实体：整列增/删/移/拷 | 📋 M2 | 自研：跨行遍历同列 `TableCell`（上游 table-column schema 为参考） |
| 图片 | 插入（PNG/JPEG）、尺寸（width/height，缺省守纵横比）、before/after 定位；M7 起 **alt 文本**（add `props.alt` 或对承载图片的段/run `set {alt}`，写 `wp:docPr/@descr`，无障碍审计读取） | ✅ M2 已实现（M7 alt 文本） | `ImagePart` + inline `Drawing`（WordHandler.Images.cs）；src 必经 workspace 沙箱；二进制导出 → 📋 余项 |
| 文本框 / 形状 | 几何、填充、线条、环绕、锚定 | 📋 M2 | DrawingML `wps:wsp` |
| 页眉 / 页脚 | 增删改，寻址 `/header[1]/p[1]`；M5 起**首页/奇偶页变体**：`/header[firstPage]`、`/header[even]`（footer 同形），add 自动接线 `w:titlePg` / `w:evenAndOddHeaders`，三变体共存、get 回报 variant | ✅ M1 已实现（M5 + 变体） | `HeaderPart` / `FooterPart` + `SectionProperties` 引用（WordHandler.HeaderFooter.cs）；read structure / render 同步暴露 |
| 节属性 | 纸张（A4/Letter/Legal…）/方向/边距 set + get（`/section[1]`），全部节枚举；M4 起**插入分节符**（`add type:sectionBreak`，kind=nextPage/continuous）+ 删除节（合并回邻节），多节文档逐节独立页面设置；M6 起**多栏排版**（columns 栏数 + columnGap 栏间距，get 回报 columns/columnGapCm） | ✅ M3 已实现（M4 分节符 + M6 分栏） | `SectionProperties`（WordHandler.Sections.cs）：`PageSize`/`PageMargin`，横竖转换守恒宽高互换；`Columns`（M6，等宽多栏）；页码格式/页面边框 → 📋 余项 |
| 分页符 / 换行符 | page break、line break、column break | ✅ M6 已实现（column break） | `Break`（`BreakValues`）：`add type:columnBreak` 写 `w:br w:type="column"`，配合多栏分节让后续文本流入下一栏；page/line break 通用糖 → 📋 余项 |
| 列表与多级编号 | 项目符号、编号、多级嵌套（level）、重新编号（listRestart）；text 视图带 `1.`/`•` 标记，HTML 渲染真 `<ol>/<ul>` | ✅ M3 已实现 | `NumberingDefinitionsPart`（WordHandler.Lists.cs）：add p 的 `list`/`level`/`listRestart` props；num/abstractNum 按需创建 |
| 超链接 | 外链（url）+ 书签锚点（anchor）行内插入 | ✅ M3 已实现 | `Hyperlink` + `HyperlinkRelationship`（WordHandler.Links.cs）；Hyperlink 样式自动接线 |
| 书签 | bookmarkStart/End 包裹目标段落，供锚点链接跳转 | ✅ M3 已实现 | `BookmarkStart`/`BookmarkEnd`（WordHandler.Bookmarks.cs）；Word 命名规则校验（字母/下划线开头、≤40 字符） |
| 批注 | 永久批注读写删 + 锚点回报；M3 起支持**线程回复**（`/comment[@id=N]` 上 `add type:reply`），`read --view comments` 按线程展示 | ✅ M2 已实现（M3 + 回复） | `WordprocessingCommentsPart`（WordHandler.Comments.cs）+ w15 commentsExtended（`w15:paraIdParent` 挂线程） |
| 脚注 / 尾注 | 脚注 + 尾注增删、`/endnote[@id=N]` 寻址、text 视图 `[^e1]` 标记与正文清单 | ✅ M3 已实现（M4 + 尾注） | `FootnotesPart`（WordHandler.Footnotes.cs）+ `EndnotesPart`（WordHandler.Endnotes.cs），separator/continuation 缺省自动补 |
| 域（field） | `add type:field`：pageNumber（PAGE）/ numPages（NUMPAGES）/ date（DATE + format picture）/ docTitle（TITLE），`leadingText` 前缀字面文本 —— 「Page X of Y」页脚 = 两个域组合；body 与页眉页脚段落均可；get/render 回报 kind + instruction；M8 起含 SEQ/REF/PAGEREF（题注/交叉引用，见下行） | ✅ M5 已实现（M8 + SEQ/REF/PAGEREF） | `SimpleField`（WordHandler.Fields.cs）+ M8 复杂域（WordHandler.Captions.cs）；其余上游域 + fieldchar/instrtext 通用低层 → 📋 余项 |
| 题注 / 交叉引用（caption / crossRef） | `add type:caption`：`label` Figure/Table/Equation（各自 SEQ 计数器）+ `text` + `position` before/after，生成 Caption 样式段落（「Figure 」+ SEQ 域 + 「: text」，书签包裹），寻址 `/caption[@label=Figure][i]`（按 label 1-based），`get` 回读 label/number/text；`add type:crossRef`：`to` 指向题注路径 + `show` `labelAndNumber`(默认)/`numberOnly`/`text`/`page`（page 用 PAGEREF），追加 REF 域 + 缓存显示文本；编号缓存（Word 打开/F9 重算）→ `caption_numbers_cached` warning | ✅ M8 已实现 | `SEQ`/`REF`/`PAGEREF` 复杂域 + `BookmarkStart/End` 包裹（WordHandler.Captions.cs）；`/caption[@label=…][i]` 双括号虚拟路径在 handler 层先于核心 grammar 拦截 |
| 目录 TOC | 插入（`add type:toc`，levels/title/position）：扫描标题样式生成超链接条目 + TOC 域指令；`/toc[1]` get（entryCount/levels/title）/remove；`read --view structure` 暴露 | ✅ M4 已实现 | `SdtBlock` 容器 + `TOC \o \h \z \u` 域 + PAGEREF 占位（WordHandler.Toc.cs）；页码无后端可算 → 写入时报 `toc_pages_unknown` warning（Word 打开/F9 时重算），诚实不估算 |
| 派生字段重算 | TOC 页码 / PAGE / NUMPAGES / 交叉引用刷新 | 🗺 M3 | 无 Word 后端可依赖：headless 渲染估算页码（上游非 Windows 同样是估算） |
| 图表 | chart + chart-axis / chart-series 子实体 | 📋 M2 | `ChartPart`（DrawingML charts，与 xlsx/pptx 共享实现） |
| 公式（equation） | OMML 数学公式读写：`add type:equation`（latex + display），行内 `/body/p[i]/omath[j]` get 回报存档 LaTeX，`read --view text` 显示 `$…$`/`$$…$$` 标记 | ✅ M6 已实现 | 自研 LaTeX → OMML 转换器（无 LaTeX 依赖，`AIOffice.Word.Equations`）：分式/根式/上下标/大型算符/矩阵（pmatrix/bmatrix/…）/`\left\right`/重音/`\text`/希腊字母/算符关系；未知命令降级为字面 run + `equation_partial` warning（文件仍校验通过）；原始 LaTeX 以 `mc:Ignorable` 厂商属性存档，round-trip 字节忠实；pptx/xlsx 公式 → 🗺 M9 |
| 水印 | 文本水印（`add type:watermark`，text/diagonal/color），写入每个页眉；`/watermark[1]` get/remove；无页眉时自动建默认页眉 | ✅ M4 已实现 | 页眉内 VML `v:shape`（WordHandler.Watermark.cs，PowerPlusWaterMarkObject 命名对齐 Word）；图片水印 → 📋 余项 |
| 表单域 | formfield 读写 + 表单字段清单导出 | 🗺 M3 | `FormFieldData` 系列 |
| 内容控件（sdt） | `add type:contentControl`：text / dropdown（items + 选中校验）/ date（ISO→yyyy-MM-dd）/ checkbox（checked），title/tag；`/sdt[@tag=X]` 或 `/sdt[i]` set 写值、`read --view fields` 清单（kind/tag/title/value/items）、remove 默认保内容；body 与页眉页脚均可 | ✅ M7 已实现 | `SdtBlock` 块级内容控件（WordHandler.ContentControls.cs）：text=`SdtContentText`、dropdown=`SdtContentDropDownList`、date=`SdtContentDate`、checkbox=w14 `SdtContentCheckBox`（自动声明 w14 mc:Ignorable 保校验）；tag 唯一；行内 sdt run / repeatingSection → 📋 余项 |
| 修订（track changes） | 文本级 ins/del 写入（`--track --author`）、accept/reject（按 `/revision[@id=N]` 或范围）、`read --view revisions`；M3 起 accept/reject 同样处理**格式修订**（w:rPrChange/w:pPrChange，reject 还原旧格式） | ✅ M2 已实现（M3 + 格式修订处理） | `InsertedRun`/`DeletedRun` + `RunPropertiesChange`/`ParagraphPropertiesChange`（WordHandler.Track.cs）；M4 起 `op:replace` + `track:true` 逐命中生成 w:del+w:ins 修订对（body 作用域）；**写入** tracked 格式变更、moveFrom/moveTo → 🗺 M3 |
| 编辑权限范围 | permStart/permEnd 区域授权 | 🗺 M3 | `PermStart` / `PermEnd` |
| OLE 对象 | 嵌入对象读写、二进制导出 | 🗺 M3 | `EmbeddedObjectPart` |
| 制表位 | tabs 简写、ptab | 📋 M2 | `Tabs` / `PositionalTab` |
| 文档默认值 | docDefaults 默认字体/字号 | 🔜 M1 | `DocDefaults`（StylesPart 内） |
| 文档设置 | autoHyphenation、evenAndOddHeaders、docGrid、CJK spacing、protection=forms | 📋 M2 | `SettingsPart`（`Settings` 子元素族） |
| RTL 与 locale | direction=rtl（段/run/表/节）、rtlGutter、按 locale 设分文种默认字体与自动 RTL | ✅ M6 已实现（段/run/表 RTL） | `BiDi`（段落 `w:bidi`，同时右对齐）/ `RightToLeftText`（run `w:rtl`）/ `BiDiVisual`（表格 `w:bidiVisual`，镜像列序）（WordFormatting.cs / WordHandler.Tables.cs）；`rtl` bool prop，get 回报；节级 rtlGutter / locale→字体映射表 → 📋 余项 |

---

## 2. xlsx 能力

| 能力 | 说明 | AIOffice 状态 | 实现备注 |
|---|---|---|---|
| 工作表管理 | 增/删/改名/移动；改名后公式引用级联更新 | ✅ M0 已实现 | ClosedXML `IXLWorksheet`（rename 自动级联引用） |
| 工作表属性 | visible/hidden/veryHidden、sheetView 基础 | 🔜 M1 | `IXLWorksheet.Visibility` |
| 打印设置 | 方向/纸张（A4/Letter…）/fitToWidth/fitToHeight/printArea set + get 反映 | ✅ M3 已实现 | `IXLPageSetup`（ExcelHandler.SheetProps.cs）；printTitleRows/Cols、手动分页符 → 📋 余项 |
| 单元格读写 | 值（类型推断：数字/日期/布尔/文本）、公式写入（`=` 自动识别） | ✅ M0 已实现 | `IXLCell.Value` / `FormulaA1` |
| 查找替换（find/replace） | 共享 M4 契约 `{op:replace, props:{find, replace, regex?, matchCase?, wholeWord?, inFormulas?}}`；作用域 = sheet / 区域 / 单元格；只匹配**文本**单元格值（数字/布尔/日期不参与），`inFormulas:true` 才进公式文本（含 `=` 形态），替换后公式经正常管线重求值；regex 走 .NET Regex + 2 秒超时（超时→invalid_args）；零命中 = ok + `find_no_match` warning；逐 op 回报 `replacements` 计数 + ≤20 个单元格路径 | ✅ M4 已实现 | 自研（ExcelHandler.Replace.cs）；替换结果一律落为字面文本（不再自动改型）；wholeWord 用前后向断言（兼容首尾非词字符的 pattern） |
| 区域寻址与批量赋值 | `/Sheet1/A1:C10`、`/Sheet1/row[3]`、`/Sheet1/col[C]`（M4）、带空格表名 `/'Q3 Data'/B2`；M4 起 2D 批量写入两种形态：**锚点式**（`set /Sheet1/A2` + `values:[[…]]`，范围按数组推断、行可参差）与**区域式**（`set /Sheet1/A2:D101`，形状必须精确匹配，错配报 invalid_args 并写明两侧形状）；类型推断与单格 set 一致（数字/布尔/日期/公式串） | ✅ M0 已实现（M4 批量 2D 写入 + M6 就地流式写入） | `IXLRange`；自研寻址解析；>50k 单元格写入**空白 sheet** 走 OpenXml SAX 流式写（ExcelBulkWrites.cs，等值律测试钉死两条路径产物内容一致）；M6：**既有大 sheet 就地流式写入**（ExcelStreamingWrite.cs，见下行「就地流式写入」） |
| 就地流式写入（M6 旗舰） | `edit … stream:true`（或 >20 MB 文件）：整批均为可流式 op（`set value` / `set values` 批量 / 公式串）时，经 SAX writer **原地改写**既有大工作簿目标 sheet part，深单元格/批量写秒级、内存有界；任一 op 非流式（bold/fill/numberFormat/merge…）则整批回退 DOM；响应标 `streamed:true` | ✅ M6 已实现 | `ExcelStreamingWrite.cs`：`TryPlan` 判定整批可流式 → SAX 流式改写，否则 DOM 回退；公式写入无缓存值（Excel 打开时重算）；50 MB+ 工作簿就地改写实测秒级 |
| 字符格式基础 | 粗/斜/字号/字色/填充色 | ✅ M0 已实现 | `IXLStyle.Font` / `Fill` |
| 数字格式 / 对齐 / 边框 | 内置与自定义 number format、对齐、边框样式 | 🔜 M1 | `IXLStyle.NumberFormat` / `Alignment` / `Border` |
| 富文本单元格 | 单元格内多 run 异格式 | 🔜 M1 | `IXLCell.GetRichText()` |
| 行列操作 | 插入（before/after）/删除/隐藏、行高（pt，0–409）列宽（字符，0–255）；列按字母寻址 `/Sheet1/col[C]`；插删后公式引用自动重写（测试断言 `=SUM` 区间随插删伸缩） | ✅ M4 已实现 | `InsertRowsAbove/Below`、`InsertColumnsBefore/After`、`Row.Height`/`Column.Width`、`Hide/Unhide`（ExcelHandler.RowsCols.cs；ClosedXML 自动调整引用）；单元格级 shift 语义（删除补位/插入挤位）→ 📋 余项 |
| 单元格 shift 语义 | 删除补位（left/up）、插入挤位（right/down），对齐 Excel UI 对话框 | 📋 M2 | `InsertCellsAfter` / `Delete(XLShiftDeletedCells)` |
| 合并单元格 | merge/拆分；query 枚举 merge/mergedrange | 🔜 M1 | `IXLRange.Merge()` / `Worksheet.MergedRanges` |
| 公式求值（回读） | 读取公式单元格时返回计算值（常用函数集） | ✅ M0 已实现 | ClosedXML CalcEngine；不支持的函数→`formula_not_evaluated` warning + 建议（绝不静默给空值） |
| 现代函数写入前缀 | XLOOKUP/FILTER/LET/LAMBDA 等写入时自动加 `_xlfn.`（FILTER 加 `_xlfn._xlws.`）、spill 引用 `A1#` | 📋 M2 | 自研 qualifier（上游 `ModernFunctionQualifier` 为行为参考）；保证 Excel 打开能算 |
| 动态数组 / 迭代求解求值 | FILTER/SORT/UNIQUE/SEQUENCE 求值；RATE/IRR 迭代解 | 🗺 M3 | 上游同样不求值（写入可用）；评估自研迭代器的性价比 |
| 排序 | 多列 `COL DIR` 列表、表头感知；拒绝含合并/公式区域；超链接/批注/条件格式随行迁移 | 📋 M2 | `IXLRange.Sort`；旁挂元数据跟随需自研 |
| 自动筛选 | 区域 AutoFilter（true 设 / false 清，单表一处） | ✅ M3 已实现 | `Range.SetAutoFilter()`（ExcelHandler.SheetProps.cs）；get sheet 反映 `autoFilter` 区域 |
| 冻结窗格 | freezeRows/freezeCols（0 清除单轴） | ✅ M3 已实现 | `SheetView.Freeze`（ExcelHandler.SheetProps.cs）；get sheet 反映 |
| 命名区域 | 定义（workbook/sheet 两种 scope）/get/remove；`/name[@name=X]` 寻址；公式直接可用（`=SUM(SalesData)` 真求值出缓存值） | ✅ M3 已实现 | `IXLWorkbook.NamedRanges`（ExcelNames.cs）；Excel 命名规则校验 |
| 表（listobject） | 区域转表、内建表样式目录、totals row | 📋 M2 | `IXLTable` + `XLTableTheme` |
| 条件格式 | cellIs（7 算子含 between）/ colorScale（2-3 色）/ dataBar / containsText 四类，`/Sheet1/conditionalFormat[i]` 寻址 | ✅ M2 已实现 | `IXLConditionalFormat`（ExcelConditionalFormats.cs）；iconset / formulacf / topN / aboveAverage / dateoccurring / duplicate / unique 七类 → 🗺 M3 |
| 数据验证 | `add type:dataValidation`（op path = 区域）：list（字面 `values` 或 `sourceRange` 引用，Excel 显示下拉）、wholeNumber/decimal/date/textLength（operator between/>/>=/</<=/==/!= + value/value2）；allowBlank、input/error 提示语、errorStyle stop/warning/information；`/Sheet1/dataValidation[i]` get/remove，structure 逐 sheet 列出 | ✅ M5 已实现 | `IXLDataValidation`（ExcelDataValidations.cs）；自定义公式验证 → 📋 余项 |
| 迷你图 | `add type:sparkline`（op path = 宿主单元格）：line / column / winLoss、dataRange、color、markers（仅 line）；`/Sheet1/sparkline[i]` get/remove，structure 列出分组 | ✅ M5 已实现 | `IXLSparklineGroups` + x14 extLst 修正（ExcelSparklines.cs，剥离 xr2:uid 保 validator 全绿） |
| 线程批注 | `add type:comment`（path = 单元格）建线程根、`add type:reply`（path = `/Sheet1/comment[@id=GUID]`）追加回复；`read --view comments` 按线程列出、get 单元格回显线程、删根整线程移除 | ✅ M5 已实现 | 自研 `xl/threadedComments` + persons part（ExcelComments.cs）+ 旧客户端 legacy note 退化镜像 |
| 图表 | bar/line/pie + **scatter/area**（M3）增删读（`/Sheet1/chart[1]` 寻址、anchor 简写）；scatter 双数值轴、首列为数值 X | ✅ M1 已实现（M3 + scatter/area） | **ClosedXML 不支持图表**：自研 OpenXml `ChartPart` + `DrawingsPart`（`ExcelCharts.cs`）；轴/系列深编辑、bubble/radar/combo → 📋 余项 |
| 图片 | 插入/锚定（anchor 单元格）/缩放（widthPx/heightPx，缺省守纵横比）；PNG/JPEG（头部嗅探） | ✅ M2 已实现 | 自研 OpenXml `DrawingsPart` oneCellAnchor（ExcelImages.cs，ClosedXML 图片 API 绕过）；SVG 双 part → 🗺 M3 |
| 批注（note） | 单元格批注增/读/删 + 作者；**寻址契约（已定）：一律经单元格** —— add `{type:note, path:/Sheet1/B2}`、remove `{path:/Sheet1/B2, props:{target:"note"}}`、get 单元格回显 `note`，`read --view structure` 列出全部带注单元格；不提供 `note[i]` 序号寻址（单元格才是稳定地址） | ✅ M4 已实现 | `IXLCell.CreateComment`/`HasComment`/`Clear(Comments)`（ExcelHandler.Notes.cs）；round-trip 法则实测：无编辑重存时 comments1.xml 与 VML part 字节不变；富文本批注样式 → 📋 余项 |
| 形状 / 文本框 | 浮动 shape/textbox；选择器需枚举 grpSp 内叶子 | 🗺 M3 | ClosedXML 不支持：OpenXml `DrawingsPart`（xdr:sp） |
| 切片器 | 表/透视表切片器（`add type:slicer`：表列 `props:{column}` 路径 `/Sheet/table[@name=X]`，或透视字段 `props:{field}` 路径 `/Sheet/pivot[i]`；可选 `caption`/`x`/`widthCells`/`heightCells`；寻址 `/Sheet1/slicer[i]` / `slicer[@name=X]`，`get`/`remove`/`read --view structure` 回读） | ✅ M8 已实现 | 全程原始 OpenXml 自研（ClosedXML 0.105 无法创建 slicer parts，仅字节级保留）：`SlicerCachePart` + 表 `x15:tableSlicerCache` / 透视 `x14:slicerCachePivotTables`，工作簿 `extLst` 引用 cache（ExcelSlicers.cs）；列/字段名大小写不敏感匹配，错名 → candidates |
| 数据透视表 | rows/columns/filters + values（sum/average/count/min/max）、targetSheet 自动建、`pivot[@name=X]` 寻址、refreshOnLoad | ✅ M2 已实现 | `IXLPivotTable`（ExcelPivots.cs）+ OpenXml 原始断言测试；layout/topN/calculatedField/showDataAs/日期分组/缓存共享 → 🗺 M3 |
| Excel 表（ListObject） | `add type:table`（op path = 区域，首行表头）：name + 内置样式（none/light1-21/medium1-28/dark1-11，如 `medium2`）+ headerRow/totalsRow + totals（列名→函数 sum/average/count/min/max/stdDev/var）+ bandedRows/Columns；结构化引用 `=SUM(Sales[Amount])` 求值（表先于保存入模型）；`/Sheet1/table[@name=Sales]` get 回报 range/style/columns/totals；remove 撤销表对象但**保留单元格数据** | ✅ M6 已实现 | `IXLRange.CreateTable` / `IXLTable`（ExcelTables.cs）；样式映射 `XLTableTheme.FromName`，汇总函数 `XLTotalsRowFunction`；未知样式/函数/列名 → invalid_args 列候选 |
| 大纲分组（outline） | `add type:group` 行/列跨段折叠分组：`/Sheet1/row[a]:row[b]` 或 `/Sheet1/col[a]:col[b]`，`collapsed:true` 折叠新组，嵌套抬升大纲级；remove 同跨段反分组一级；行/列 get 与 `read --view structure` 回报 outlineLevel/collapsed | ✅ M6 已实现 | `IXLRows.Group` / `Collapse` / `Ungroup`（ExcelGroups.cs）；`row[a]:row[b]`/`col[a]:col[b]` 跨段在 Core DocPath 新增 ElementSpan 形式解析，CLI/MCP 命令面直达 |
| 超链接 | 单元格超链接读写删：`hyperlink` prop（`https://…` 外链 / `#Sheet!A1` 内链）+ `hyperlinkTooltip`；get 回显，空串清除 | ✅ M5 已实现 | `IXLCell.SetHyperlink`（ExcelHandler.Edit.cs cell props 路径） |
| 命名单元格样式 | `add type:cellStyle path:/styles`（name + numberFormat/bold/italic/fill/color/border 任意子集）一次定义，`set {cellStyle:"X"}` 套用到单元格/区域，`read --view styles` 清单 + `get /style[@name=X]`、remove 删除；套用时把具体格式直绘到目标单元格（缓存显示即时反映），保存后再补真 `cellStyle` 表项让 Excel 样式库可见 | ✅ M7 已实现 | ClosedXML 0.105 无命名样式表 API：样式定义存于工作簿 custom property `_aioffice_cellStyles`（JSON map，逃过 ClosedXML 重存），套用经 ClosedXML 直写格式 + 保存后原始 `cellStyle`/`cellStyleXfs` 补丁（ExcelCellStyles.cs，validator 全绿）；`/style[@name=X]` 寻址 |
| 工作簿级属性 | password、calc.mode=manual、calc.refMode=r1c1 | 📋 M2 | `Protect` / `CalculateMode` / OpenXml `WorkbookProperties` |
| CSV/TSV 导入 | `create --from data.csv`（见跨格式「CSV/TSV 导入 xlsx」行）；`read --view csv [--sheet][--range]` 反向导出 | ✅ M5 已实现 | 自研解析 + ClosedXML/SAX 写入（ExcelCsv.cs / ExcelHandler.Csv.cs）；stdin、表头→AutoFilter+冻结、起始单元格 → 📋 余项 |
| 行按列名查询 | `row[Salary>5000]` 式按表头列名过滤 | 📋 M2 | 自研：表头行映射 + 选择器扩展 |
| 大工作簿流式读取 | 超过 20 MB（或 `stream:true`）自动走 SAX：`read --view stats/text`、单元格/区域 `get` 不加载 DOM、按需提前停 | ✅ M3 已实现 | `OpenXmlPartReader` 原始扫描（ExcelStreaming.cs）；41 MB / 33 万行 stats ≈ 2 秒；流式**写**通道：批量 2D 写入空白 sheet（M4，ExcelBulkWrites.cs）、csv 导入（M5，同一 SAX 路径），M6 起**既有大 sheet 就地流式写入**（见区域寻址行「就地流式写入」，整批可流式时不再 ClosedXML 全量加载） |
| OLE 对象 | 嵌入对象、preview 图 | 🗺 M3 | OpenXml `EmbeddedObjectPart` |
| RTL | sheetView rightToLeft、cell/comment direction | 🗺 M9 | OpenXml `SheetView.RightToLeft`（M6 RTL 只覆盖 docx；xlsx sheetView RTL 顺延 M9） |
| 拼音指引（phonetic） | 东亚文本 phonetic runs | 🗺 M3 | OpenXml `PhoneticRun` |

---

## 3. pptx 能力

| 能力 | 说明 | AIOffice 状态 | 实现备注 |
|---|---|---|---|
| 幻灯片增删改移 | 新建（基于默认版式或 `props:{layout:N}` 指定版式）、删除、重排 | ✅ M0 已实现（M6 指定版式） | `SlidePart` + `SlideIdList` 维护；M6 `add type:slide props:{layout:N}` 套用克隆版式 |
| 幻灯片分节（section） | `add type:section`（path = `/` 根，name + afterSlide 0-based 起始）把幻灯片分组到大纲节；`set /section[i] {name}` 改名、`remove /section[i]` 删节（片留存）；`read --view outline` 按节分组、`get /section[i]` 回报名与所属片；按 sldId 跟踪，抗重排 | ✅ M6 已实现 | `p14:sectionLst`（presentation.xml extLst，`{521415D9-…}` 扩展，PptxSections.cs）；新增 `/section[i]` 寻址与 `/` 根 op 路由 |
| 幻灯片尺寸 / 宽高比 | `set /` `{slideSize:"16:9\|4:3\|16:10\|A4\|letter"}` 命名预设或显式 `{width,height}`（cm/in/emu）；改写 `p:sldSz`，既有形状坐标不变；`get /` 回报 slideSize/widthCm/heightCm/slideCount/sectionCount | ✅ M6 已实现 | `PptxSlideSize.cs`：预设 EMU 表 + `SlideSizeValues`（4:3=Screen4x3，其余 Custom）；新增 `/` 演示文稿根 get/set 路由 |
| 隐藏幻灯片 | show=false | 🔜 M1 | `Slide.Show` |
| 幻灯片克隆 | 复制 slide 含全部跨 part 关系（图片/图表/媒体） | 📋 M2 | `AddPart` + 关系 id 重映射 |
| 形状 / 文本框 | 新建、位置/尺寸（EMU 及 cm/pt/px 单位换算）、文本读写 | ✅ M0 已实现 | `P.Shape` + `a:txBody`；`Transform2D` |
| 文本段落与 run 格式 | 字体/字号/颜色/粗斜/对齐/项目符号基础 | ✅ M0 已实现 | `a:p` / `a:r` / `a:rPr`（M0 覆盖基础属性） |
| 标题占位符填充 | 新片设标题（create --title、slide title 属性） | ✅ M0 已实现 | 版式 placeholder 匹配（phType=title/ctrTitle） |
| 任意占位符 | phType=body/subTitle/footer/dt/sldNum 等定位与填充 | 🔜 M1 | `PlaceholderShape` 索引/类型匹配 |
| 形状外观基础 | 纯色填充、线条颜色、预设几何（rect/roundRect/ellipse/triangle/diamond/arrow + line 连接线）、翻转（flipH/flipV） | ✅ M3 已实现 | `a:prstGeom` + `P.ConnectionShape`（PptxEditor.cs）；渐变/效果/旋转 → 📋 余项 |
| 形状外观高级 | 渐变/图片填充、阴影/发光效果、旋转、图片亮度对比度 | 📋 M2 | `a:gradFill` / `a:effectLst` / `xfrm@rot` |
| effective 继承值溯源 | 读出属性的最终生效值及来源（shape→layout→master 链） | 🗺 M3 | 自研继承解析器（上游 effective.X.src 为行为参考） |
| 母版 / 版式寻址 | `/master[1]`、`/master[1]/layout[2]` 路径，get/query | ✅ M1 已实现（M6 转可编辑） | `SlideMasterPart` / `SlideLayoutPart` 枚举；M6 起编辑落地（见下行「母版/版式编辑」），不再 `unsupported_feature` |
| 母版 / 版式编辑 | 对 master/layout 上元素 add/set/remove | ✅ M6 已实现 | `PptxMasters.cs`：`set /master[m]`（background + accent1..6 主题强调色）、`set /master[m]/layout[l]`（background）、`add /master[m] {type:layout, basedOn?}`（克隆既有版式）、`/master[m]/shape[i]` 与 `/master[m]/layout[l]/shape[i]` 复用 slide 形状 op；克隆版式可经 `add type:slide props:{layout:N}` 套用到新片；无引用版式才可 remove |
| 图片 | 插入（PNG/JPEG）/位置/尺寸（单位换算，缺省守纵横比）/name；返回稳定 `shape[@id=N]` 路径；M7 起 **alt 文本/标题**（`set {altText}`/`{altTitle}`，写 `p:nvPicPr/cNvPr/@descr`/`@title`，无障碍审计读取，SVG 渲染出 `<title>`） | ✅ M2 已实现（M7 alt 文本） | `ImagePart` + `P.Picture`（PptxImages.cs）；src 必经沙箱；SVG/裁剪/滤镜 → 🗺 M9 |
| 表格 | `add type:table`（rows×cols、headerRow、x/y/w、columnWidths）+ light/medium/dark 三档直绘样式（header 填充 + 斑马行，不依赖主题）；单元格 text/`mergeRight`/`mergeDown`、加行（`add type:row`）；`/slide[i]/table[k]/tr[r]/tc[c]` 寻址，get 回报 rowsDetail + 合并形状，SVG 渲染画出真网格 | ✅ M5 已实现 | `a:tbl`（PptxTables.cs，graphicFrame 手写）；内建 tableStyles 目录引用、列实体操作 → 📋 余项 |
| 图表基础 | 柱/线/饼增删读（`/slide[i]/chart[k]` 寻址）；跨文档 `dataFrom` 直接取 xlsx 数据；M4 起新建图表**默认嵌入工作簿** + 系列/分类写 Sheet1 引用（`c:numRef`/`c:strRef` + 缓存），PowerPoint 右键「编辑数据」直接可用（get 报 `dataEditable:true`）；旧式纯缓存图表 `set {embedData:true}` 一键退化改造 | ✅ M3 已实现（M4 + 嵌入工作簿） | 自研 `ChartPart` + `EmbeddedPackagePart`（PptxCharts.cs/PptxChartWorkbook.cs，c:externalData autoUpdate=false）+ 命令层 dataFrom 展开；轴/系列子实体 → 📋 余项 |
| 图表高级 | pieOfPie/barOfPie、轴线/网格线逐属性、dropLines/hiLowLines/upDownBars、图表动画（chartBuild） | 🗺 M3 | ChartPart 深水区 |
| 连接线 | from/to 锚接形状（接受 @name）、起止端点改绑 | 📋 M2 | `P.ConnectionShape` + `a:stCxn`/`a:endCxn` |
| 组合（group） | 建组/解组、get/query 深度遍历、link/tooltip | 📋 M2 | `P.GroupShape` 递归 |
| 备注（notes） | 读写删每页演讲者备注（`/slide[i]/notes` set/add/remove/get，多段落） | ✅ M2 已实现 | `NotesSlidePart` + notes master 自动接线（PptxNotes.cs） |
| 幻灯片背景 | slide 级纯色背景（真 `p:bg`，非全幅矩形） | ✅ M2 已实现 | `P.Background` + `a:solidFill`（set/add `background` prop）；渐变/图片背景 → 🗺 M3 |
| 批注（legacy） | 经典批注增/读/删（`add type:comment`，text/author/x/y），`/slide[i]/comment[@id=N]` 稳定寻址，`read --view structure` 列出；M5 起**回复线程**：`add type:reply`（path = 父批注），`read --view comments` 按线程嵌套 `replies`，删根连带回复 | ✅ M4 已实现（M5 + 回复线程） | `SlideCommentsPart` + `CommentAuthorsPart`（PptxComments.cs，作者去重复用）；回复 = p:cm + p15:threadingInfo/parentCm（PowerPoint 2013+ 显示真线程） |
| 现代线程批注 | p188 modern comments（线程/回复）round-trip | 🗺 M3 | `powerPointComments` part（OpenXml 3.x 已建模） |
| 动画 | 进入预设 appear/fade/flyIn(8 方向)/wipe(4 方向)；M5 起**强调** pulse/grow/spin/colorPulse（colorPulse 收 color）与**退出** fadeOut/flyOut/wipeOut（退出尾帧挂 hide set）；click/withPrevious/afterPrevious 触发、duration/delay；`/slide[i]/animation[k]` 寻址增删，structure 视图按播放顺序列出（含 class=entrance/emphasis/exit）；M6 起**时间轴重排**（`move /slide[i]/animation[2] before/after /slide[i]/animation[1]`）+ `set` 重调既有动画 props | ✅ M4 已实现（M5 强调/退出 + M6 时间轴编辑） | `p:timing` 树手写（PptxAnimations.cs + 预设类映射，OpenXml 无高层 API）；M6 时间轴 move/set 落地；其余上游预设（15 emphasis + 16 exit 全集）、多效果链、motion path、repeat/restart/autoReverse、动画预设++ → 🗺 M9 |
| 切换（transition） | none/fade/push/wipe + 时长（`transitionDuration`，p14:dur 毫秒），get/outline 回读 | ✅ M3 已实现 | `P.Transition`（PptxTransitions.cs）；p15 12 种扩展预设 → 📋 余项 |
| Morph 平滑切换 | p14 morph + `!!` 命名匹配感知 | 🗺 M3 | p14 扩展 + 对象名约定 |
| 媒体（video/audio） | 嵌入/链接、loop、autoStart | 🗺 M3 | `VideoReferenceRelationship` + `p:nvPr` media |
| 公式（equation） | OMML 公式形状 | 🗺 M9 | `a14:m` 内嵌 Math（M6 公式已落地 docx；pptx 公式形状顺延 M9，复用 `AIOffice.Word.Equations` 转换器） |
| zoom | summary / section / slide zoom 跳转对象 | 🗺 M3 | p15 `sectionZoom`/`slideZoom` |
| OLE 对象 | 嵌入对象 + preview 图 | 🗺 M3 | `graphicFrame` + `p:oleObj` |
| 3D 模型 | glTF 模型嵌入、rotation=ax,ay,az | 🗺 M3 | am3d 扩展 part |
| SmartArt | **只读**落地（M5）：`read --view structure` 逐片列出 smartArt、`get /slide[i]/smartart[k]` 输出连接序嵌套节点树（text/children）、宿主 frame 回指 smartArtPath、query 可命中；编辑报 `unsupported_feature` + workaround | ✅ M5 已实现（只读） | `DiagramDataPart` 走查（PptxSmartArt.cs）；写入（data/layout/colors part 组）→ 🗺 深水区 |
| 超链接 / 动作 | 形状点击动作 `set {hyperlink}`：外链 url / 幻灯片跳转 `#slide:N` / 放映动作 `#first`·`#last`·`#next`·`#prev`·`#end`，空串清除；文本超链接 `set {linkText}`（同目标语法包裹 runs）；`get` 回读规范形式（url / `#slide:N` / `#first…`） | ✅ M8 已实现 | `a:hlinkClick`（url → `r:id` 关系、跳转 → `ppaction://hlinksldjump`、放映 → `ppaction://hlinkshowjump?jump=…`）（PptxHyperlinks.cs）；tooltip → 📋 余项 |
| 主题 / 文档级属性 | theme 颜色字体、defaultFont、show.loop、print.what | 📋 M2 | `ThemePart` / `PresentationPropertiesPart` |
| RTL | shape/notes direction、图表轴 RTL | 🗺 M9 | `a:bodyPr@rtlCol` / `a:pPr@rtl`（M6 RTL 只覆盖 docx；pptx 形状 RTL 顺延 M9） |
| 形状 z 序 | `move` 到 front/back/forward/backward，重排 spTree 绘制顺序 | ✅ M3 已实现 | spTree 子节点重排（PptxEditor.cs）；front = 最后绘制（最上层） |
| 逐页渲染 SVG/HTML | 每页 slide 输出 SVG（或 HTML 退化；M3 起含迷你图表渲染、预设几何、线条） | ✅ M0 已实现 | 自研渲染器（基础形状/文本；覆盖面随元素能力增长） |
| 截图网格 / 页范围 | PNG 缩略图网格（--grid）、html 页范围 | 🔜 M1 | 单页 PNG 已落地（`render --to png --scope /slide[N]`）；网格拼接与页范围未实现 |

---

## 4. 统计

| 范围 | ✅ 已实现（M0–M9） | 🔜 M1 余项 | 📋 M2 余项 | 🗺 1.0 之后 | ❌ 不计划 | 合计 |
|---|---|---|---|---|---|---|
| 跨格式通用 | 30 | 3 | 5 | 2 | 3 | 43 |
| docx | 24 | 1 | 6 | 5 | 0 | 36 |
| xlsx | 26 | 4 | 6 | 5 | 0 | 41 |
| pptx | 21 | 3 | 5 | 10 | 0 | 39 |
| **合计** | **101** | **11** | **22** | **22** | **3** | **159** |

> 解读：M0 的 31 项里有相当一部分是 AI 原生层（envelope、快照、rev、sandbox、schema、MCP）—— 这是 AIOffice 的差异化地基，上游完全没有。格式能力本身（docx 4 / xlsx 5 / pptx 5）在 M0 刻意收窄到"段落、单元格、形状"三个最小核心，换取每一项都带测试、过 `OpenXmlValidator`、守住 round-trip 字节一致律。
>
> M1（v0.2.0）新增 6 项：docx 页眉/页脚、xlsx 图表（bar/line/pie，自研 ChartPart）、pptx 母版/版式只读寻址、跨格式 PNG 渲染（headless 浏览器）、实时预览 server、浏览器选区回读（`data-aio-path` 契约，原排 M2 提前落地）。
>
> M2（v0.3.0）新增 11 项（含 2 个新行）：docx 修订（文本级 ins/del + accept/reject）、批注、样式管理、图片；xlsx 数据透视表、条件格式（4 类）、图片；pptx 真背景、演讲者备注、图片；跨格式大文件守卫（`file_too_large` + `AIOFFICE_MAX_FILE_MB`）。计划中的"大文件流式"没有按 M2 交付——它需要一轮专门的基准驱动打磨，移入 M3；M2 以尺寸守卫诚实兜底。
>
> M3（v0.4.0，功能第一指令）新增 16 项（含 3 个新行：跨文档 dataFrom、大工作簿流式读取、形状 z 序）：docx 列表/超链接/书签/脚注/节属性 5 项翻绿，批注与修订两行就地深化（线程回复、格式修订 accept/reject）；xlsx 流式读取、打印设置、自动筛选、冻结窗格、命名区域 5 项翻绿，图表行就地深化（scatter/area）；pptx 图表、切换、预设几何、z 序 4 项翻绿；跨格式 render PDF 与 dataFrom 落地，大文件守卫默认翻转为不限大小（`AIOFFICE_MAX_FILE_MB` 改 opt-in）。
>
> M4（v0.5.0）新增 8 项翻绿（含 1 个新行：xlsx 查找替换）：跨格式**全文 find/replace** 落地（三 handler 共享一份契约 + 命令层 `"/"` 文档级展开 + CLI `--find/--replace` 糖，docx `--track` 出修订对）；docx **TOC、水印** 2 项翻绿，脚注/尾注（+尾注）、节属性（+插入分节符）、修订（+tracked replace）3 行就地深化；xlsx 行列操作、单元格批注 2 项翻绿，区域寻址行深化（批量 2D 写入 + 空白 sheet SAX 流式写）；pptx **进入动画、经典批注** 2 项翻绿，图表行深化（嵌入式工作簿 → PowerPoint「编辑数据」可用）。工程面：tag 触发的 GitHub release 自动化（6 平台单文件二进制 + SHA256SUMS）。
>
> M5（v0.6.0，markdown/csv 桥）新增 12 项翻绿（含 2 个新行：markdown 桥、xlsx 线程批注）：跨格式 **markdown↔docx 桥**（Markdig 导入 + GFM 导出 round-trip）与 **csv↔xlsx 桥**（`create --from` / `read --view csv`，前导零保文本）落地，`IFormatHandler` 增量获得 `CreateFrom` 默认钩子（无导入器格式诚实拒绝）；docx **表格格式、单元格合并、域** 3 项翻绿，页眉页脚行深化（firstPage/even 变体）；xlsx **数据验证、迷你图、超链接、CSV 导入、线程批注** 5 项翻绿；pptx **原生表格、SmartArt（只读）** 2 项翻绿，动画行深化（强调/退出预设）、批注行深化（回复线程）。其余项按余项标记排队（M6 窗口：就地写入流式、插件机制、公式 OMML、RTL 深化、现代 pptx 批注 part、动画时间轴编辑），每一项落地时同步更新本表状态与"实现备注"中的真实 API。
>
> M6（v0.7.0，深水区攻坚）新增 9 项翻绿（含 5 个新行：xlsx 就地流式写入 / Excel 表 / 大纲分组、pptx 幻灯片分节 / 幻灯片尺寸）：docx **公式（LaTeX → OMML 自研转换器）、右到左/双向、多栏分节（含 columnBreak）** 3 项翻绿；xlsx **就地流式写入**（既有大工作簿 SAX 原地改写）、**Excel 表（ListObject + 汇总 + 结构化引用求值）**、**行/列大纲分组** 3 项翻绿；pptx **母版/版式编辑（背景 + 强调色 + 克隆版式）、幻灯片分节、幻灯片尺寸/宽高比** 3 项翻绿，动画行深化（时间轴 move/set 重排）。Core 新增 `DocPath` 寻址形式：`/`（演示文稿根，pptx 幻灯片尺寸 + 分节 + 文档级 get）与 `row[a]:row[b]`/`col[a]:col[b]` 元素跨段（xlsx 大纲分组）—— 这是把 M6 能力接到真实 CLI/MCP 命令面的关键一步（之前这些路径在 `EditOp.ParseBatch` → `DocPath.Parse` 网关被拒，仅直构 EditOp 的格式层测试可达）。
>
> M7（v0.8.0，交付前审计）新增翻绿（含 4 个新行：跨格式**审计**、跨格式**文档属性**、docx**内容控件**、xlsx**命名单元格样式**）：三格式共享 **`audit` 动词 + `office_audit` MCP 工具（第 15 个工具，MCP 表面从 14 → 15 工具）** —— 无障碍 + 质量 lint，findings 是数据（`ok:true`/exit 0），`--fix` 只做安全自动修复（占位 alt、表头行、文档/幻灯片标题、孤儿书签）后重审返回 `{fixed, remaining}`；docx **文档属性（core + 类型化 custom）、内容控件（text/dropdown/date/checkbox + `read --view fields`）** 翻绿，图片行深化（alt 文本）；xlsx **命名单元格样式（`add type:cellStyle` + `read --view styles`）、文档属性** 翻绿；pptx **alt-text/标题/阅读顺序审计 + 显式 `altText`/`altTitle`、文档属性** 翻绿。Core 增量：`IAuditor` 接口 + `AuditFinding`/`AuditResult`/`AuditSummary`/`AuditOptions` records（三 handler 各实现，命令层共享 `AuditVerb`）；`office_read` view 枚举新增 `properties`/`fields`/`styles`。深水区余项顺延 M8/M9（语义 diff 动词已于 M8 交付，见下；pptx/xlsx 公式、跨格式 docx↔pptx 转换、OLE 对象、插件机制、动画预设++ 顺延 M9）。

> M8（v0.9.0，对比与审阅）新增翻绿（含 4 个新行：跨格式**语义对比 diff**、docx**题注/交叉引用**、xlsx**切片器**、pptx**形状超链接/动作**）：三格式共享 **`diff` 动词 + `office_diff` MCP 工具（第 16 个工具，MCP 表面从 15 → 16 工具）** —— 语义对比当前文档与基线（同格式文件 `other` 或自身快照 `--snapshot N`，跨格式 → `invalid_args`），返回**已排序、跨平台确定性**变更清单 `{changes:[{kind:added\|removed\|modified\|moved, path, before?, after?, detail?}], summary, baseline}`（按 `(path, kind)` 序贯排序，LCS/内容哈希区分 moved），变更是数据（`ok:true`/exit 0），`--view summary\|detailed`，快照 diff 工作流（`diff f --snapshot 1` = 上次编辑变更集）；docx **题注（Figure/Table/Equation + SEQ）/交叉引用（REF/PAGEREF + show 模式）** 翻绿，寻址 `/caption[@label=…][i]`；xlsx **切片器（表列 / 透视字段，原始 OpenXml 自研）** 翻绿，寻址 `/Sheet1/slicer[i]`；pptx **形状超链接/动作（url / `#slide:N` / `#first…#end` 放映动作 / linkText）** 翻绿。Core 增量：`IDiffer` 接口 + `DiffChange`/`DiffResult`/`DiffSummary` records（`DiffResult.FromChanges` 序贯排序，三 handler 各实现，命令层共享 `DiffVerb`）。深水区余项整体顺延 M9（跨格式 docx↔pptx 转换已于 M9 交付，见下；pptx/xlsx 公式复用 `AIOffice.Word.Equations`、OLE 对象、插件机制、API 稳定性承诺、动画预设++ 顺延 1.0 之后）。

> M9（v0.10.0，pre-1.0 capstone）新增翻绿（1 个新行：跨格式**互转 convert**，并统一既有 md/csv 桥与 pdf 渲染入口）：三格式共享 **`convert` 动词 + `office_convert` MCP 工具（第 17 个工具，MCP 表面从 16 → 17 工具）** —— 跨格式内容互转，按 (源扩展名, 目标扩展名) 路由：同族文本桥 **docx↔md**（markdown 桥）/ **xlsx↔csv**（csv 桥）复用既有代码，跨格式经**内容中立模型**（`INeutralConvertible`：源 `ExportNeutral` → 目标 `ImportNeutral`）覆盖 **docx↔pptx / docx↔xlsx / pptx↔xlsx / 任意↔md**（md 经命令层 `NeutralMarkdown` 序列化器，故任意 office 格式可转 md），**any→pdf/png/svg/html/txt** 经 render 层；转换跨格式本质有损 —— 丢失项（动画/图表/SmartArt/精确样式/公式取值/markdown 不支持的颜色与下划线）汇入单条 `convert_lossy` warning + `data.dropped`，目标**新建**（覆盖既有 → `convert_overwrite` warning），同扩展名 → `invalid_args`（「use edit」），未知目标 → `unsupported_feature`（列出受支持目标）；envelope `data:{from,to,blocksWritten,dropped:[…],written}`。Core 增量：`NeutralDoc`/`NeutralBlock`/`NeutralRun`/`ImportResult` records + `INeutralConvertible` 接口（与 M7 `IAuditor`、M8 `IDiffer` 同模式：Core 定义、三 handler 各实现）；命令层 `ConvertVerb` + `NeutralMarkdown`（CLI `convert` 与 MCP `office_convert` 同产物）。**1.0 加固**：每个 CLI 动词均有 help 主题并出现在 `schema`（新增 `convert` help 主题：(src→dest) 矩阵 + 有损说明 + 示例）；`doctor` / `office_status` 新增 **capabilities 自省块**（verbs 数 / MCP 工具数 / 支持格式 / convert 源与目标 / render 目标 / audit 类别），一次调用即可自省全表面。**Toward 1.0** 余项：pptx/xlsx 公式（经共享转换器）、OLE 对象、插件机制、API 稳定性承诺。
