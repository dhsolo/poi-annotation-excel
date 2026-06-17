# poi-annotation-excel 功能清单

> 基于 Apache POI 5.5.1 的**注解驱动** Excel 导出/导入库（JDK 17）。坐标 `io.github.dhsolo:poi-annotation-excel`，基础包 `io.github.dhsolo.poi.excel`。
> 本文是功能总览/索引；每个功能的可运行示例见 [`USAGE.md`](USAGE.md)。

---

## 0. 统一门面 `ExcelUtil`

| 方向 | 方法 | 说明 |
|------|------|------|
| 导出 | `toBytes` / `toInputStream` / `export(out)` / `exportLocal(path)` | 产出 `byte[]` / 流 / 写出到流 / 写本地文件 |
| 导出 | `upload(uploader, name, ...)` | 交给自定义 `ExcelUploader`（S3/MinIO/OSS 等）|
| 导入 | `importExcel(in, Class)` | 注解模型 → `List<T>`（自动解析 `@ExcelInfo` 布局：跳序号列/except 列）|
| 导入 | `importExcel(in, sheetIndex, startRow, Class, columns...)` | 指定 sheet/起始行/显式列 |
| 导入 | `importExcelToMap(in, columns...)` | 无需 POJO，返回 `List<Map>` |
| 导入 | `importExcel(in, Class, ExcelReadListener)` | 逐行回调，错误不中断（`onRow`/`onError`/`onFinish`）|

**三种导出入口**：① 注解模型对象；② `ExcelCreatorBuilder`（动态、无模型）；③ 预构建的 `ExcelCreator`（多个=多 Sheet）。
**格式**：导出 `xlsx`/`xls`；导入自动识别 `xlsx`/`xls`/`csv`（按文件魔数）。

---

## 1. 注解清单（共 18 个）

### 类级
| 注解 | 作用 |
|------|------|
| `@ExcelInfo` | 整张 Sheet 配置：`sheetName`、`excelType`、`isBigData`(流式)、`needOrder`/`orderColumnSpan`(序号列)、`startRow`/`exceptColumnNum`(导入)、`titleHeight`/`headerHeight`、`validateRowCount`、`imageReadTimeOut`/`imageSeparator`/`imageResize`、`noneCellDefaultValue` |

### 列与表头
| 注解 | 作用 |
|------|------|
| `@ExcelColumn` | 列定义：`columnName`/`index`/`columnWidth`/`nullable`/`mergeCellIndex`(横向合并)/`needMergeCell`(纵向合并)/`noneCellDefaultValue`；`translate`(内联翻译)/`sourceField`/`sourcePath`(取值重定向)；**`groups`(多级表头路径)** |
| `@ExcelColumnParent` | 两行分组合并表头（父标题跨子列，子列用 `sourceField`/`sourcePath`）|
| `@ExcelTitle` | 顶部标题行（跨所有列合并）|
| `@ExcelContext` | 在 `BEFORE_TITLE`/`BETWEEN_TITLE_HEADER`/`BETWEEN_HEADER_DATA`/`AFTER_DATA` 注入上下文行（字段或方法）|
| `@ExcelRow` + `@ExcelCell` | 声明式自定义行，逐格指定内容（`value` 字面量 / `field` 取字段）|
| `@ExcelData` | 标记数据列表字段（每个元素=一行）|
| `@ExcelInfoChild` | 把嵌套对象的列**拍平**进当前表（多层，点路径如 `addr.geo.zip`，含类型环检测）|

### 取值/翻译/校验
| 注解 | 作用 |
|------|------|
| `@ExcelDateFormat` | 日期/时间格式 pattern（导出格式化、导入回解析）|
| `@ExcelImage` | 图片列（`imageDownPath`/`imageVisitPrev`）|
| `@ExcelImageResize` | 全局图片缩放 |
| `@ExcelListBox` | 下拉列表（字段固定值 / 方法动态值，`isNeedAddTranslationException`）|
| `@ExcelTranslateMethod` | 方法级值翻译（码→显示文本，**仅导出**；导入的值翻译用 `@ExcelColumn(translate=...)` 内联映射）|
| `@ExcelCustomValidateMethod` | 自定义行校验（拿到整行 `ExcelRowData`）|
| `@ExcelFormula` | 单元格公式，列引用用 `@Column(列名)` 语法 |
| `@ExcelHeaderColumnMapping` | 导入时方法返回 `Map<表头文本,字段名>`，类导入据此读表头行**按表头名重排列**（列序变化/表头≠字段名时用；未映射列跳过）|

---

## 2. 表头能力

- **标题**：`@ExcelTitle`（跨列合并居中）。
- **列头**：`@ExcelColumn(columnName, index)`，按 `index` 排序。
- **两行分组表头**：`@ExcelColumnParent`（父标题横跨子列；非分组列纵向合并跨两行）。
- **多级表头（3 层及以上）**：`@ExcelColumn(groups = {"上级","次级"...})` —— 任意深度，相邻同前缀横向合并、短路径叶子纵向合并到底；与 `@ExcelColumnParent` 互斥（混用 fail-fast）。
- **上下文行**：`@ExcelContext` / `@ExcelRow`+`@ExcelCell` / `creator.addDiyRowContext(...)`。
- **序号列**：`needOrder` + `orderColumnSpan`（横向合并跨度，纵向跨全表头）。
- **列合并**：`mergeCellIndex`（一列占多格）、`needMergeCell`（相邻同值纵向合并）。
- **自动列宽**（抽样估宽、CJK 记 2 宽，公式取缓存结果）、**冻结窗格**、**显式列宽/行高**。

---

## 3. 类型化写值与精度

- `Number`/`Boolean` 写**原生类型**（可排序/求和）。
- 日期：`Date`/`Calendar`/`LocalDate`/`LocalDateTime` → 类型化日期单元格 + pattern；`LocalTime` 按一天的分数写。
- **大数精度保护**：`Long`/`BigInteger` 超 2^53、`BigDecimal` 有效位 >15 → 按文本写，避免 double 丢精度（如 19 位 ID）。

---

## 4. 图片

- **导出**（`@ExcelImage`）：远程 URL **并行下载**（共享线程池）、按字节嗅探格式**原样嵌入**（PNG 透明保留）、`STORED` 直接注入 `xl/media`（不重压缩）、缩放、多图按分隔符逐列排列、坏图跳过不中断。
- **SSRF 策略**（`ImageDownloadPolicy`）：协议白名单、私网/回环阻断（可选）、单图 64 MB 上限、超时、安全模式下不跟随重定向。
- **导入**：从单元格锚点提取图片落盘并回填访问 URL。
- **模板图片占位**：见第 8 节。

---

## 5. 数据校验 / 下拉 / 级联

- **下拉**：`@ExcelListBox`（固定值或方法动态值）。
- **级联联动**：`CascadeValidateModelBuilder`，**任意层级**（INDIRECT + 隐藏 sheet + 定义名，自动处理 Excel 非法字符/大小写冲突）。
- **公式校验**：`@ExcelFormula`，列引用 `@Column(列名)`。
- **自定义校验**：`@ExcelCustomValidateMethod`，回调拿到整行数据（可转 Bean）。
- 校验覆盖行数 `validateRowCount`（默认 1000）。

---

## 6. 多 Sheet / 复杂多区 / 嵌套

- **多 Sheet**：多个 `ExcelCreator`/`ExcelCreatorBuilder` = 多个工作表。
- **复杂多区（单 Sheet 多区段）**：模型实现 `ComplexExcelModel`，`getComplexModels()` 返回多个子区段，纵向堆叠在同一 Sheet（支持区段用多级表头）。
- **子表样式保留**：跨簿克隆样式，子表自定义样式不丢。

---

## 7. 大数据 / 流式

- 导出 `isBigData=true`（默认）→ SXSSF 滑窗刷盘，内存平稳。
- 导入 `StreamingExcelReader`（SAX）→ 低内存只读、前向、按 `@ExcelColumn` 位置或**表头名**绑定 Bean。

---

## 8. 模板填充 `ExcelTemplateFiller`

- `${占位}` 标量替换；`fill`/`fillAll`/`fillBean`。
- 列表区 `${list.xxx}`：`fillList`(Map) / `fillListBeans`(POJO)，自动扩行并下移下方图片锚点。
- **图片占位** `${@image:key}` / `${list.@image:key}`：值可为 `byte[]`/`File`/`InputStream`/`String` URL（走同套 SSRF 守卫下载），多图按分隔符逐列排列。
- 实现 `AutoCloseable`（自建 workbook 用 try-with-resources 自动关）。

---

## 9. 导入细节

- **DOM** `ExcelImportor`：随机访问、图片、自定义校验；扁平列按 `sourcePath` 唯一键映射（同名子字段不串值）。
- **赋值**：先 setter，再回退直接字段赋值；坏单元格只丢**该行**不中断整次导入。
- **类型转换**：整数族（Int/Long/Short/Byte/BigInteger）统一截断；日期优先用列自身 `@ExcelDateFormat` pattern 回解析。
- **布局对齐**：类导入自动按 `@ExcelInfo` 的 `needOrder`/`orderColumnSpan` 跳序号列、`exceptColumnNum` 跳列。
- **CSV**：作为 POI `Sheet`/`Row`/`Cell` 适配层接入同一导入管线（字符集自动探测）。

---

## 10. 样式与输出

- 三套样式：标题 / 表头 / 数据，均可自定义（`CellStyleManager`）。
- 输出：`byte[]` / `InputStream` / `OutputStream` / 本地文件 / 自定义 `ExcelUploader`。
- 资源安全：导出器/模板器幂等、临时文件按生命周期清理、共享线程池在全部实例关闭时优雅 shutdown。

---

## 速查：我该用哪个？

| 场景 | 用法 |
|------|------|
| 有模型类、规整导出 | `@ExcelInfo`+`@ExcelColumn` → `ExcelUtil.toBytes(model)` |
| 无模型、列运行时才知道 | `ExcelCreatorBuilder.create(...).columns(...).data(...)` |
| 三级以上复杂表头 | `@ExcelColumn(groups = {...})` |
| 百万行、内存敏感（导出/导入）| `isBigData=true` / `StreamingExcelReader` |
| 改现成模板填值/填图 | `ExcelTemplateFiller` |
| 联动下拉 | `CascadeValidateModelBuilder` |
| 单 Sheet 多张表 | `ComplexExcelModel` |
| 导入列顺序会变 | `@ExcelHeaderColumnMapping`（DOM）或流式 `StreamingExcelReader.readAsBeans(..., matchByHeaderName=true)` |
