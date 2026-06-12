# goldens/

基准产物（golden files），由格式包的 xunit 测试消费。Golden reference files
consumed by the format packages' xunit tests.

- **round-trip law / 往返定律**：对这里的每个夹具文件，打开后不做任何编辑直接
  保存，所有 zip part 必须字节不变；若某个 part 合理地发生变化，对应测试必须
  注明确切原因。
- **validator oracle**：每个变更类测试结束后，OpenXmlValidator 必须报告 0 错误。
- 子目录按格式划分：`goldens/docx/`、`goldens/xlsx/`、`goldens/pptx/`
  （由各格式包的测试按需创建）。
- 渲染基准（render 输出的 html/svg 快照）同样放在对应格式子目录下，文件名与
  来源夹具一致。

不要手工编辑这里的二进制文件；用 `aioffice` 重新生成并在 PR 中说明原因。
Do not hand-edit binaries here; regenerate them with `aioffice` and explain why
in the PR.
