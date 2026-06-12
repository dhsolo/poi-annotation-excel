# poi-annotation-excel 使用手册

基于 Apache POI 的注解驱动 Excel 导出/导入库。本手册覆盖全部公开用法与每个配置项。

- 坐标：`io.github.dhsolo:poi-annotation-excel`
- 基础包：`io.github.dhsolo.poi.excel`
- 运行环境：JDK 17+，POI 5.2.5（传递依赖）

---

## 目录

1. [快速上手](#1-快速上手)
2. [注解参考（全属性）](#2-注解参考全属性)
3. [导出 API](#3-导出-api)
4. [动态导出（Builder）](#4-动态导出builder)
5. [自定义行 / 标题 / 序号 / 冻结 / 列宽 / 样式](#5-自定义行--标题--序号--冻结--列宽--样式)
6. [多 Sheet 与嵌套列](#6-多-sheet-与嵌套列)
7. [图片导出](#7-图片导出)
8. [导入 API](#8-导入-api)
9. [低内存流式导入 StreamingExcelReader](#9-低内存流式导入-streamingexcelreader)
10. [级联下拉 CascadeValidateModelBuilder](#10-级联下拉-cascadevalidatemodelbuilder)
11. [下拉校验 / 翻译 / 自定义校验 / 公式](#11-下拉校验--翻译--自定义校验--公式)
12. [模板填充 ExcelTemplateFiller](#12-模板填充-exceltemplatefiller)
13. [类型化写值规则](#13-类型化写值规则)
14. [安全（图片下载）](#14-安全图片下载)
15. [枚举与常量](#15-枚举与常量)
16. [限制与注意事项](#16-限制与注意事项)

---

## 1. 快速上手

### 导出（注解模型）

```java
@ExcelInfo(sheetName = "设备列表")
public class DeviceExportModel {
    @ExcelTitle  private String title;            // 可选标题行
    @ExcelData   private List<Device> data;       // 必填：数据行来源

    @ExcelColumn(columnName = "设备名称", index = 1) private String devName;
    @ExcelColumn(columnName = "状态", index = 2, translate = {"0:离线", "1:在线"}) private Integer status;
    // getter/setter（模型需要无参构造器）
}

// 列字段（devName/status...）按 name 从每个 Device 行对象读取（需对应 getter）
ExcelUtil.export(response.getOutputStream(), "设备列表.xlsx",
        new DeviceExportModel().setTitle("月度报表").setData(deviceList));
```

### 导入（同一个注解模型）

```java
List<DeviceExportModel> rows = ExcelUtil.importExcel(inputStream, DeviceExportModel.class);
```

---

## 2. 注解参考（全属性）

### `@ExcelInfo`（类级，描述整张 Sheet）

| 属性 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `sheetName` | String | `""` | Sheet 名（空则用默认名） |
| `excelType` | String | `"xlsx"` | 输出类型，`ExcelCreator.XLSX` / `ExcelCreator.XLS` |
| `needOrder` | boolean | `false` | 是否在最前面加自增「序号」列 |
| `orderColumnSpan` | int | `1` | 序号列横向跨/合并的列数（`needOrder=true` 时有效） |
| `titleHeight` | int | `2000` | 标题行高（POI 单位 = 1/20 磅） |
| `headerHeight` | int | `2000` | 表头行高 |
| `startRow` | int | `0` | **导入**时数据起始行（0 基） |
| `sheetNum` | int[] | `{0}` | **导入**时读取的 Sheet 索引 |
| `exceptColumnNum` | int[] | `{}` | **导入**时跳过的列索引 |
| `imageReadTimeOut` | int | `500` | 远程图片下载超时（ms） |
| `imageSeparator` | String | `""` | 单元格内多图 URL 的分隔符（空则按 `,`） |
| `pictureInnerType` | int | `0` | 图片锚定方式，见 [AnchorType](#15-枚举与常量) |
| `isBigData` | boolean | `true` | 是否用 SXSSF 流式写（大数据省内存）。**注意默认即流式**：行窗口外不可回访，`getWorkBook()` 返回 SXSSF 实例；小表或需后处理时显式置 `false` 用全内存 XSSF |
| `validateRowCount` | int | `1000` | 下拉校验/`@ExcelFormula` 公式预填覆盖的数据行数（自首个数据行起算） |
| `imageResize` | `@ExcelImageResize` | 不缩放 | 全局图片缩放策略 |
| `noneCellDefaultValue` | String | `""` | 字段为 null/空时写入的默认文本 |

```java
// 导出：大数据流式 + 序号列 + 默认值
@ExcelInfo(sheetName = "设备列表", isBigData = true,
           needOrder = true, orderColumnSpan = 1,
           noneCellDefaultValue = "-")
public class DeviceExportModel { /* ... */ }

// 导入：读第 2 个 Sheet、从第 3 行开始、跳过第 0 列
@ExcelInfo(sheetNum = {1}, startRow = 2, exceptColumnNum = {0})
public class DeviceImportModel { /* ... */ }
```

### `@ExcelData`（字段）
标记承载数据行的字段，类型为 `List<行对象>`。每个 Sheet 必填一个。

### `@ExcelTitle`（字段或方法）
提供表头之上的标题文本，须为 `String`。可选。

### `@ExcelColumn`（字段，定义一列）

| 属性 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `columnName` | String | `""` | 列标题（空则用字段名） |
| `index` | int | 必填 | 列顺序（升序排列；导入按位置绑定） |
| `nullable` | boolean | `true` | **导入**时该列是否允许为空 |
| `columnWidth` | int | `20` | 列宽（字符数） |
| `needMergeCell` | boolean | `false` | 相邻同值是否纵向合并 |
| `mergeCellIndex` | int | `1` | 合并跨度 |
| `noneCellDefaultValue` | String | `""` | 该列 null 时的默认值（覆盖 Sheet 级） |
| `translate` | String[] | `{}` | 内联翻译，`{"0:否","1:是"}`（导出 key→显示；导入反向） |
| `sourcePath` | String | `""` | 点路径取嵌套值，如 `"device.type"` |
| `sourceField` | String | `""` | 同级字段重定向，如显示列读另一个字段 |

```java
@ExcelInfo(sheetName = "设备")
public class DeviceModel {
    @ExcelData private List<Device> data;

    // 1) 相邻同值纵向合并（如同一区域多行只显示一次）
    @ExcelColumn(columnName = "区域", index = 1, needMergeCell = true)
    private String area;

    // 2) sourcePath：从行对象的嵌套属性取值（row.getType().getName()）
    @ExcelColumn(columnName = "型号", index = 2, sourcePath = "type.name")
    private Object typeName;          // 仅作锚点，值来自 type.name

    // 3) sourceField：本列显示另一字段的值（导出 displayName，读 rawName）
    @ExcelColumn(columnName = "名称", index = 3, sourceField = "rawName")
    private Object displayName;

    // 4) 导入允许为空 + 列宽
    @ExcelColumn(columnName = "备注", index = 4, nullable = true, columnWidth = 30)
    private String remark;
}
```

### `@ExcelDateFormat(pattern)`（字段）
标记日期列并指定格式，默认 `"yyyy-MM-dd HH:mm:ss"`。导出时该列按**类型化日期单元格**写入并套用该格式。

```java
@ExcelColumn(columnName = "创建日期", index = 1)
@ExcelDateFormat(pattern = "yyyy-MM-dd")
private LocalDate createdDate;

@ExcelColumn(columnName = "下单时间", index = 2)
@ExcelDateFormat(pattern = "yyyy/MM/dd HH:mm")
private LocalDateTime orderTime;

@ExcelColumn(columnName = "班次", index = 3)
@ExcelDateFormat(pattern = "HH:mm:ss")
private LocalTime shiftTime;        // 时间单元格（一天的分数）
```

### `@ExcelImage`（字段）+ `@ExcelImageResize`
见 [图片导出](#7-图片导出)。
- `@ExcelImage.imageVisitPrev`：URL 前缀（拼到字段值前）。
- `@ExcelImage.imageDownPath`：本地缓存目录。
- `@ExcelImageResize.needResize/resizeWidth/resizeHeight`：缩放开关与尺寸（默认 500×500）。

### `@ExcelListBox`（字段或方法）
给列附加下拉校验列表。`listTextBox` 为静态候选；标注在方法上则运行时动态提供（方法名由 `columnName` 关联到列）。`isNeedAddTranslationException` 控制是否为翻译值加例外。

下拉校验默认覆盖自首个数据行起的 **1000 行**，可通过 `@ExcelInfo(validateRowCount = ...)`、
`ExcelCreator.setValidationRowCount(int)` 或 Builder 的 `validationRowCount(int)` 调整
（模板预期接收更多行时调大，小模板可调小瘦身）。`@ExcelFormula` 的公式预填行数同受此配置控制。

### `@ExcelTranslateMethod(columnName)`（方法）
为指定列提供自定义翻译函数，方法须返回 `Function`。比 `@ExcelColumn.translate` 更灵活（复杂映射）。同时存在时方法优先。

### `@ExcelCustomValidateMethod(columnName)`（方法）
为**导入**指定列提供自定义校验。标注的方法须返回 `ExcelCustomValidate`（`boolean validate(ExcelRowData)` + `String errorMessage()`），`columnName` 填该列对应的**字段名**。也可用编程式 `ExcelModel.setExcelCustomValidate(...)`（手动列场景）。

### `@ExcelFormula(formula)`（字段）
让该列每个数据单元格写入公式。公式中用 `@Column(列名)` 引用其它列（按列标题匹配），框架按行展开为「Excel 列号 + 当前行号」。

### `@ExcelRow`（字段或方法）+ `@ExcelCell`
向 Sheet 注入一行**自定义（非数据）行**，跨所有列合并。
- `@ExcelRow.order`：多条自定义行的顺序（多条时**须唯一**，否则顺序不确定）。
- `@ExcelRow.cells`：`@ExcelCell[]`，逐格指定内容（`value` 字面量 / `field` 取字段值，二选一）。

```java
@ExcelInfo(sheetName = "报表")
public class ReportModel {
    @ExcelData private List<Row> data;
    @ExcelColumn(columnName = "状态", index = 1) private String status;

    // 形式 1：不写 cells —— 用字段自身的值作整行内容（整行合并）
    @ExcelRow(order = 0)
    private String footnote = "脚注：仅供参考";

    // 形式 2：cells = @ExcelCell(value=...) —— 整行写静态文本
    @ExcelRow(order = 1, cells = @ExcelCell(value = "数据来源：XX 系统"))
    private String source;

    // 形式 3：cells = @ExcelCell(field=...) —— 整行取另一字段的值
    @ExcelRow(order = 2, cells = @ExcelCell(field = "projectName"))
    private String projectRow;
    private String projectName = "智能运维平台";
}
```

### `@ExcelInfoChild`（字段）
把嵌套子对象的 `@ExcelColumn` 列**拍平**进父 Sheet。子列按各自 `@ExcelColumn.index()` 与父列**统一排序**。
- **导出**：通过路径 `父字段名.子字段名` 读取行对象上的嵌套值（行对象需有 `get父字段().get子字段()`）。
- **导入**：自动重建嵌套对象并填充（父字段类型需有无参构造器，子字段需 setter）。
- 子列支持 `columnName/index/translate/日期格式/图片/公式`；级联与方法级下拉/翻译不适用于子列。导入按各列的 `sourcePath` 唯一键映射，**不同父对象的同名子字段（如 `customer.name` 与 `supplier.name`）不会串值**。支持**多层嵌套**（子对象内再放 `@ExcelInfoChild`，路径如 `addr.geo.zip`；有类型环检测）。

```java
@ExcelInfo(sheetName = "订单")
public class OrderModel {
    @ExcelData private List<OrderRow> data;
    @ExcelColumn(columnName = "订单号", index = 1) private String orderNo;
    @ExcelInfoChild private Customer customer;   // Customer 的列被拍平进来
}
public class Customer {
    @ExcelColumn(columnName = "客户名", index = 2) private String name;   // 读 row.getCustomer().getName()
    @ExcelColumn(columnName = "电话",  index = 3) private String phone;
}
```

多层嵌套（子对象内再放 `@ExcelInfoChild`，叶子列路径如 `addr.geo.zip`）：
```java
@ExcelColumn(columnName = "单号", index = 1) private String no;
@ExcelInfoChild private Address addr;          // → addr.*

public class Address {
    @ExcelColumn(columnName = "城市", index = 2) private String city;   // 路径 addr.city
    @ExcelInfoChild private Geo geo;                                     // → addr.geo.*
}
public class Geo {
    @ExcelColumn(columnName = "邮编", index = 3) private String zip;    // 路径 addr.geo.zip
}
```

### `@ExcelColumnParent(value, columns)`（字段）
生成**两行合并表头**：上行父标题 `value`（横向合并跨子列），下行各子列标题。子列为 `@ExcelColumn` 数组，**每个子列须用 `sourceField`（或 `sourcePath`）指定取值来源**（子列无对应 Java 字段）。子列与其它列一起按 `index()` 统一排序（同组子列需相邻索引）。

```java
@ExcelColumn(columnName = "名称", index = 1) private String name;

@ExcelColumnParent(value = "销售额", columns = {
    @ExcelColumn(columnName = "Q1", index = 2, sourceField = "q1"),
    @ExcelColumn(columnName = "Q2", index = 3, sourceField = "q2")
})
private Object sales;   // 锚点字段；值来自行对象的 getQ1()/getQ2()
```
> 非分组列（含序号列）会纵向合并跨两行表头、标题居中显示（无序号列或单列序号时）；仅导出。

### `@ExcelContext`（字段或方法）
在 Sheet 指定位置注入额外的上下文行。`type` 取 `ContextType`：`BEFORE_TITLE`（标题前）、`BETWEEN_TITLE_HEADER`（标题与表头之间）、`BETWEEN_HEADER_DATA`（表头与数据之间）、`AFTER_DATA`（数据之后）。

```java
@ExcelInfo(sheetName = "月报")
public class MonthlyModel {
    @ExcelTitle private String title;
    @ExcelData  private List<Row> data;

    @ExcelContext(type = ContextType.BETWEEN_TITLE_HEADER)
    private String subTitle = "统计周期：2024-06";

    @ExcelContext(type = ContextType.AFTER_DATA)
    public String footer() { return "导出时间：" + LocalDate.now(); }   // 方法形式
}
```

### `@ExcelHeaderColumnMapping`（方法）
**导入**时提供动态「表头文本 → 字段名」映射（无参方法返回 `Map<String,String>`）。当 Excel 表头文本与字段不一一对应、或表头可能变化时使用。

```java
@ExcelInfo(sheetName = "导入")
public class ImportModel {
    private String productName;
    private java.math.BigDecimal unitPrice;
    // setters ...

    @ExcelHeaderColumnMapping
    public Map<String, String> headerMapping() {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("Product Name", "productName");   // 表头文本 → 字段名
        map.put("Unit Price",   "unitPrice");
        return map;
    }
}
```

---

## 3. 导出 API

全部通过门面 `ExcelUtil`（注解模型）：

```java
// 写到输出流（name 为文件名，无需扩展名）
ExcelUtil.export(OutputStream out, String name, Object excelModel);

// 导出前追加自定义配置（自定义行、冻结等）
ExcelUtil.export(OutputStream out, String name, Object excelModel, Consumer<ExcelCreator> configurator);

// 写到本地文件
ExcelUtil.exportLocal(String localPath, Object excelModel);

// 转字节 / 输入流
byte[]      ExcelUtil.toBytes(Object excelModel);
InputStream ExcelUtil.toInputStream(Object excelModel);
InputStream ExcelUtil.toInputStream(Object excelModel, Consumer<ExcelCreator> configurator);

// 上传（自定义上传策略 ExcelUploader<R>）
<R> R ExcelUtil.upload(ExcelUploader<R> uploader, String name, Object excelModel);
<R> R ExcelUtil.upload(ExcelUploader<R> uploader, String name, Object excelModel, Consumer<ExcelCreator> cfg);
```

`configurator` 用法（在 `createExcel()` 之前调整）：

```java
ExcelUtil.export(out, "report.xlsx", model, creator -> {
    creator.addDiyRowContext("数据来源：XX 系统", true);   // 合并的说明行
    creator.setFreeze(0, 2, 0, 0);                        // 冻结前 2 行
});
```

---

## 4. 动态导出（Builder）

无模型类时用 `ExcelCreatorBuilder`（适合动态列）：

```java
ExcelUtil.export(out, "report.xlsx",
    ExcelCreatorBuilder.create("Sheet1")
        .title("销售报表")
        .data(salesList)                                // List<Bean> 或 List<Map>
        .columns("姓名:name", "金额:amount", "日期:date") // "标题:字段名"
        .autoSizeColumns(true));
```

`ExcelCreatorBuilder` 全部方法：

| 方法 | 说明 |
|------|------|
| `create(sheetName)` | 静态工厂 |
| `excelType(String)` / `excelType(ExcelType)` | 输出类型 |
| `bigData(boolean)` | SXSSF 流式 |
| `title(String)` | 标题行 |
| `header(String[])` | 直接给表头数组 |
| `data(Object)` | 数据（List<Bean>/List<Map>） |
| `needOrderNum(boolean)` / `orderColumnSpan(int)` | 序号列 |
| `rowHeight/titleRowHeight/headerRowHeight(int)` | 行高 |
| `columns(String... "标题:字段")` | 列定义（简写） |
| `columns(String[] headers, String[] fields)` | 列定义（分开给） |
| `addColumn(int index, ExcelModel)` / `columnMapping(Map)` | 精细列定义 |
| `columnMerge(Map)` / `addMerge(int col, String field)` | 列合并 |
| `columnWidth(int col, int width)` | 列宽 |
| `autoSizeColumns(boolean)` | 抽样自动列宽 |
| `pictureType(int)` / `anchorType(AnchorType)` | 图片锚定 |
| `imageReadTimeOut(int)` / `imagesSeparator(String)` | 图片下载 |
| `noneCellDefaultValue(String)` | 空值默认文本 |
| `freezeRow(int)` / `freezeColumn(int)` / `freeze(int rows,int cols)` | 冻结窗格 |
| `titleCellStyle/cellStyle/headerCellStyle(CellStyle)` | 自定义样式 |
| `child(LinkedList)` / `addChild(ExcelCreator)` | 子 Sheet |
| `build()` / `toBytes()` / `toInputStream()` | 产出 |

也可一次导出多个 Builder（多 Sheet）：

```java
ExcelUtil.export(out, "multi.xlsx", builder1, builder2);
byte[] bytes = ExcelUtil.toBytes(builder1, builder2);
```

---

## 5. 自定义行 / 标题 / 序号 / 冻结 / 列宽 / 样式

### 序号列
`@ExcelInfo(needOrder = true, orderColumnSpan = 1)`。`orderColumnSpan > 1` 时「序号」横向合并若干列。

### 自定义行（DIY 上下文行）
`ExcelCreator.addDiyRowContext(...)`（通过 configurator 调用），重载：

```java
creator.addDiyRowContext("说明文字");                       // 一行
creator.addDiyRowContext("说明文字", true);                  // 是否整行合并
creator.addDiyRowContext("说明文字", true, 600);             // + 行高
creator.addDiyRowContext(ctx, sc, ec, cellStyleEnum, cellStyle, isMerge[, height][, afterTitle]);
```

也可在模型上用 `@ExcelRow` / `@ExcelCell` 声明式注入自定义行（见注解参考）。

### 冻结窗格
```java
creator.setFreezeRow(2);            // 冻结前 2 行
creator.setFreezeColumn(1);         // 冻结前 1 列
creator.setFreeze(sc, sr, lc, er);  // 精确指定
// Builder: .freezeRow(2) / .freezeColumn(1) / .freeze(2,1)
```

### 列宽 / 自动列宽
```java
creator.setColumnWidth(0, 30 * 256);   // POI 单位 = 1/256 字符宽
creator.setAutoSizeColumns(true);      // 抽样估算（不逐格字体度量；大表友好）
// 注解：@ExcelColumn(columnWidth = 20)
```
> 公式列按其**缓存计算结果**估宽，长公式（如 `SUM(A1:A100)`）不会把列撑宽。

### 样式
默认提供标题/表头/数据三套样式（Arial、居中、细边框）。可用 Builder 的 `cellStyle/headerCellStyle/titleCellStyle` 覆盖，或列上用 `CellStyleEnum`。

---

## 6. 多 Sheet 与嵌套列

### 多 Builder 多 Sheet
每个 Builder 一个 Sheet，一次导出：
```java
ExcelCreatorBuilder s1 = ExcelCreatorBuilder.create("设备").columns("名称:name").data(devices);
ExcelCreatorBuilder s2 = ExcelCreatorBuilder.create("告警").columns("级别:level").data(alarms);

ExcelUtil.export(out, "multi.xlsx", s1, s2);
byte[] bytes = ExcelUtil.toBytes(s1, s2);
```

### 子表保留自定义样式
子表（独立工作簿构建）并入父工作簿时，其自定义标题/表头/数据样式会被克隆保留：
```java
ExcelCreator child = ExcelCreatorBuilder.create("子表")
        .columns("名称:name").data(rows).build();
CellStyle yellow = child.createCellStyle();          // 针对子表工作簿创建
yellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
yellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
child.setCellStyle(yellow);                           // 自定义数据样式

LinkedList<ExcelCreator> children = new LinkedList<>();
children.add(child);
byte[] bytes = ExcelCreatorBuilder.create("主表")
        .columns("编号:id").data(mainRows)
        .child(children)                              // 并入为第二个 Sheet
        .toBytes();                                   // 子表数据单元格仍是黄色填充
```

### 复杂/父子 Sheet（同一 Sheet 内多张明细表）
模型实现 `ComplexExcelModel`，`getComplexModels()` 返回若干子模型（各自也是合法的 `@ExcelInfo` 模型），它们在同一 Sheet 内依次向下排布；父表的图片处理器、行号在子表间共享：
```java
@ExcelInfo(sheetName = "汇总")
public class SummaryModel implements ComplexExcelModel {
    @ExcelData private List<Head> data;
    @ExcelColumn(columnName = "汇总项", index = 1) private String item;

    @Override public List getComplexModels() {
        return List.of(new DetailModelA(detailA), new DetailModelB(detailB));
    }
}
```

### 嵌套列 / 合并表头
- **`@ExcelInfoChild`**：把嵌套子对象的列拍平进父行；**支持多层嵌套**（路径如 `addr.geo.zip`），导入按路径重建各级对象。见 [注解参考](#excelinfochild字段)。
- **`@ExcelColumnParent`**：形成两行合并表头（父列覆盖多子列）。见 [注解参考](#excelcolumnparentvalue-columns字段)。

---

## 7. 图片导出

字段同时标注 `@ExcelColumn` + `@ExcelImage`，字段值为图片 URL 或本地路径：

```java
@ExcelColumn(columnName = "照片", index = 3)
@ExcelImage(imageVisitPrev = "https://cdn.example.com/",  // URL 前缀（拼到字段值前）
            imageDownPath = "/tmp/excel-img")            // 可选：本地缓存目录
@ExcelImageResize(needResize = true, resizeWidth = 120, resizeHeight = 120)  // 字段级缩放
private String photoUrl;                                  // 值如 "a/b/photo.png" → 拼成完整 URL 下载
```

- **多图单元格**：`@ExcelInfo(imageSeparator = ",")`，字段值用分隔符拼多个 URL，自动横向展开列。
- **缩放**：`@ExcelInfo(imageResize = @ExcelImageResize(needResize = true, resizeWidth = 200, resizeHeight = 200))`。
- **锚定方式**：`@ExcelInfo(pictureInnerType = ExcelCreator.MOVE_AND_RESIZE)` 或 Builder `.anchorType(AnchorType.DONT_MOVE_AND_RESIZE)`。
- **超时**：`@ExcelInfo(imageReadTimeOut = 500)`。
- **格式保留**：源为 PNG/JPEG/GIF 时按 URL 扩展名保留原格式（含 PNG 透明）；无需 resize 且格式匹配时原始字节直传，内存友好。
- **大图量**：图片并行下载、磁盘暂存、以 STORED 注入 ZIP，避免字节堆积 OOM。

安全见 [第 14 节](#14-安全图片下载)。

---

## 8. 导入 API

```java
// 注解驱动（clazz 带 @ExcelInfo/@ExcelColumn，按列位置绑定）
List<T> ExcelUtil.importExcel(InputStream in, Class<T> clazz, ExcelModel... columns);
List<T> ExcelUtil.importExcel(InputStream in, int sheetIndex, Class<T> clazz, ExcelModel... columns);
List<T> ExcelUtil.importExcel(InputStream in, int sheetIndex, int startRow, Class<T> clazz, ExcelModel... columns);

// 手动列（无注解，按位置）
List<User> users = ExcelUtil.importExcel(in, User.class,
        ExcelModel.of("name"), ExcelModel.of("age"), ExcelModel.of("dept"));

// 跳过前 2 行元数据：
List<User> users = ExcelUtil.importExcel(in, 0, 2, User.class, ExcelModel.of("name"), ExcelModel.of("age"));

// 不要 POJO，直接取 Map：
List<Map<String,Object>> rows = ExcelUtil.importExcelToMap(in, ExcelModel.of("col1"), ExcelModel.of("col2"));
List<Map<String,Object>> rows = ExcelUtil.importExcelToMap(in, sheetIndex, startRow, columns...);
```

`ExcelModel.of(...)` 工厂：`of(field)`、`of(field, columnName)`、`of(field, translateMap)`；构造器还支持 `nullable`、`pattern`(日期)、图片等。

```java
Map<Object, Object> enabledMap = new HashMap<>();
enabledMap.put("是", 1);                                     // 单元格文本 → 写入值
enabledMap.put("否", 0);

List<User> users = ExcelUtil.importExcel(in, User.class,
    ExcelModel.of("name"),                                   // 仅字段名
    ExcelModel.of("dept", "部门"),                           // 字段名 + 列标题
    ExcelModel.of("enabled", enabledMap),                    // 翻译映射
    ExcelModel.of("age").setExcelCustomValidate(ageRule));   // 链式追加自定义校验
```

### 监听器流式（DOM）

```java
ExcelUtil.importExcel(in, DeviceImportModel.class, new ExcelReadListener() {
    @Override public void onRow(Map<String,Object> row, int rowIndex) { service.save(convert(row)); }
    @Override public void onError(String message, int rowIndex) { log.warn("第{}行: {}", rowIndex, message); }
    @Override public void onFinish(int totalRows) { log.info("共 {} 行", totalRows); }
});
// 手动列版本：ExcelUtil.importExcel(in, listener, ExcelModel.of("name"), ExcelModel.of("amount"));
```
> 注意：该监听器底层仍是 POI DOM（整本进内存）。超大文件请用第 9 节的流式读取器。

> **导入目标类要求**：无参构造器 + 各列字段的 **public setter**（框架按字段名 `setXxx(类型)` 注入；缺 setter 的列会被静默跳过、保持默认值）。

### 支持的导入类型
`String`、全部基本类型及其包装类、`BigDecimal`、`BigInteger`、`java.util.Date`、`java.sql.Date`、`java.sql.Timestamp`、`java.time.LocalDate`、`java.time.LocalDateTime`。数值容忍千分位（`"1,234,567"`）；日期容忍九种常见格式（ISO-8601、`yyyy-MM-dd HH:mm:ss`、`yyyy/MM/dd`、`yyyyMMdd` 等）。

---

## 9. 低内存流式导入 StreamingExcelReader

真·SAX 流式读取 `.xlsx`，内存有界。位于 `io.github.dhsolo.poi.excel.importor`。

```java
// 逐行回调（不累积）
StreamingExcelReader.read(in, 0, (rowIndex, cells) -> process(cells));         // cells: List<String>
StreamingExcelReader.read(in, 0, startRowInclusive, (i, cells) -> ...);        // 跳过前置行

// 表头映射 Map（首行为表头）
List<Map<String,String>> rows = StreamingExcelReader.readAsMaps(in, 0);
List<Map<String,String>> rows = StreamingExcelReader.readAsMaps(in, 0, headerRowIndex);

// 映射到注解模型（按 @ExcelColumn 位置 或 表头名 绑定，首行表头）
List<T> beans = StreamingExcelReader.readAsBeans(in, 0, T.class);
List<T> beans = StreamingExcelReader.readAsBeans(in, 0, T.class, /*按表头名*/ true);
List<T> beans = StreamingExcelReader.readAsBeans(in, 0, T.class, true, /*headerRowIndex*/ 2);
```

- `readAsBeans` 支持 `@ExcelColumn.translate` 反向映射；转换 `String`/数值/布尔/`BigDecimal`/日期时间（Excel 序列日期与文本日期均可）。
- 作用域：仅 `.xlsx`、纯表格、前向只读；公式取缓存值；不含图片/合并区/随机访问。需要这些请用 `ExcelImportor`。

---

## 10. 级联下拉 CascadeValidateModelBuilder

任意层级联动下拉（大类 → 小类 → 设备类型），EasyExcel 无等价能力。

```java
CascadeValidateModel model = CascadeValidateModelBuilder.builder("devBigType")
    .addItem("模拟量设备",
        CascadeValidateModelBuilder.builder("devSmallType")
            .addItem("电流", CascadeValidateModelBuilder.builder("pointTypeCode").addItem("CT1").addItem("CT2"))
            .addItem("电压"))
    .addItem("数字量设备")     // 叶子
    .build();
```

方法：`builder(field)` / `addItem(value)` / `addItem(value, 子builder...)` / `addItems(集合, 取名函数)` / `needAddTranslationException(boolean)` / `hasItems()` / `build()`。空安全：`addItem(name, hasChildren ? childBuilder : null)`。

从领域对象集合批量构建（`addItems` + 取名函数）：
```java
CascadeValidateModel model = CascadeValidateModelBuilder.builder("city")
    .addItems(provinces, Province::getName)   // 用每个 Province 的 name 作为候选项
    .build();
```

把级联模型挂到列上：在模型方法用 `@ExcelListBox(columnName="...")` 返回该级联，或通过模型装配。

级联的实现基于 Excel 名称管理器 + `INDIRECT`，名称由父级选项值拼接而成，受 Excel 命名规则约束：
- 选项值中的**非法字符**（空格、`-`、`/`、括号等，即字母/数字/`.`/`_` 之外的字符）会被自动替换为 `_`，
  并在下拉公式中用 `SUBSTITUTE` 镜像同样的替换——常见的真实数据可直接使用，无需预处理；
- 拼接后**超 255 字符**、**形似单元格引用**（如 `A1`）或**空值**无法替换补救，会抛出带具体选项链的
  `ExcelException`（fail-fast，不产出坏文件）；
- 名称大小写不敏感，仅大小写不同的父值（`ABC`/`abc`）会自动加后缀消歧。

---

## 11. 下拉校验 / 翻译 / 自定义校验 / 公式

### 静态下拉
```java
@ExcelColumn(columnName = "状态", index = 1)
@ExcelListBox(listTextBox = {"在线", "离线"})
private String status;
```

### 动态下拉（方法）
```java
@ExcelListBox(columnName = "city")
public String[] cityOptions() { return cityService.allNames().toArray(new String[0]); }
```

### 内联翻译 vs 方法翻译
```java
@ExcelColumn(translate = {"0:否", "1:是"})          // 简单映射
private Integer enabled;

@ExcelTranslateMethod(columnName = "level")          // 复杂映射，返回 Function
public Function<Object,Object> levelTranslate() { return v -> LEVELS.getOrDefault(v, "未知"); }
```

### 导入自定义校验
注解式（`columnName` 填字段名，方法返回 `ExcelCustomValidate`）：
```java
@ExcelColumn(columnName = "年龄", index = 1) private Integer age;

@ExcelCustomValidateMethod(columnName = "age")
public ExcelCustomValidate<Row> ageRule() {
    return new ExcelCustomValidate<>() {
        public boolean validate(ExcelRowData<Row> d) { return ((Integer) d.currentValue()) >= 0; }
        public String errorMessage() { return "年龄不能为负"; }
    };
}
```
或编程式（手动列）：`ExcelModel.of("age").setExcelCustomValidate(validator)`。

### 公式列
公式用 `@Column(列名)` 引用其它列（按列标题匹配），框架在每个数据行展开为该列的 Excel 列号 + 行号：
```java
@ExcelColumn(columnName = "合计", index = 4)
@ExcelFormula(formula = "@Column(单价)*@Column(数量)")   // 每个数据行写入：如 B2*C2、B3*C3 ...
private String total;
```

---

## 12. 模板填充 ExcelTemplateFiller

> **图片说明**：模板中已有的静态图片会原样保留；列表行扩展时，列表区**下方**的图片锚点会随插入的行数一起下移（不会压到数据上）。
>
> **列表区域说明**：每个 listKey 只识别**第一处**含 `${listKey.*}` 占位符的模板行；同一 key 在多处出现时，后续占位行不会展开。需要多个列表区域时请使用不同的 listKey。


对已有模板做 `${占位符}` 替换（不依赖注解）：

```java
ExcelTemplateFiller.of(templateInputStream)          // of(InputStream) / of(byte[]) / of(Workbook)
    .fill("reportDate", "2024-06-01")                // 单个占位符
    .fillAll(Map.of("k1", v1, "k2", v2))             // 批量
    .fillBean(reportHeader)                          // 用 Bean 的字段填充同名占位符
    .fillList("list", rowsAsMaps)                    // 列表区：${list.no} ${list.name}
    .fillListBeans("list", rowBeans)                 // 列表区（POJO）
    .fillPicture("logo", logoBytes)                  // 图片占位：${@image:logo}（byte[]/File/InputStream）
    .writeTo(response.getOutputStream());            // 或 .toBytes()
```

模板里：标量占位 `${reportDate}`；列表占位 `${list.no}`、`${list.name}`（同一行作为模板行，按列表展开）。

### 图片占位符

- **标量图片** `${@image:logo}`：配合 `fillPicture("logo", 图片)` 注册，支持 `byte[]` / `File` / `InputStream`（流会被读完但不关闭）/ **`String` URL 或本地路径**。填充时清空该单元格文本，图片以**覆盖该单元格的双锚点**插入（随行列移动缩放）；占位格在列表区下方时随扩行一起下移。
- **列表行图片** `${list.@image:photo}`：图片值直接放在行数据 Map 的 `photo` 键里（同样支持 `byte[]` / `File` / `InputStream` / `String` URL；`fillListBeans` 的 `byte[]`/`String` 字段天然可用），每行各插一张。
- **URL 下载**：`String` 值按 `http`/`https` URL 下载（非 `http` 开头按本地路径读），走与图片导出**同一套守卫**——协议白名单、`ImageDownloadPolicy`（SSRF，见第 14 节）、读超时（`imageReadTimeOut(毫秒)` 链式设置，默认 2000ms）、单图 64 MB 上限。**同一 URL 只下载一次**（失败也只试一次），同一图片多处锚定共享**一个媒体部件**（不撑大文件）。下载失败仅清空占位 + WARN，不中断填充。
- 图片**格式按字节嗅探**（PNG/JPEG/GIF/BMP）原样嵌入，不转码（PNG 透明度保留）。未注册的 key、`null` 值或无法识别的字节：仅清空占位文本、跳过插图并打 WARN 日志，不中断填充。
- 同一单元格可混排图片与文本占位：`${@image:logo}${title}` 会插图并保留 `title` 的文本替换。
- 图片大小由锚点决定（铺满占位单元格），需要更大的展示区域时请在模板里调大该行高/列宽。

---

## 13. 类型化写值规则

导出时 `setCellValue` 按运行时类型写入：

| 值类型 | 写法 |
|--------|------|
| `Number`（含包装类、BigDecimal） | 数值单元格（可排序/求和） |
| `Boolean` | 布尔单元格 |
| `Date`/`Calendar`/`LocalDate`/`LocalDateTime` | **类型化日期单元格** + 日期格式样式（`@ExcelDateFormat` 的 pattern；无则按类型默认 `yyyy-MM-dd` / `yyyy-MM-dd HH:mm:ss`） |
| `LocalTime` | **类型化时间单元格**（按「一天的分数」写入数值 + 时间格式；默认 `HH:mm:ss`，亚秒精度不保留） |
| 大整数 `Long`/`BigInteger`（超 2^53）、高精度 `BigDecimal`（有效位 >15） | **按文本写**，避免 double 精度丢失（如 19 位 ID）；代价是该格不是数值类型 |
| `null` | 空（或列/Sheet 的 `noneCellDefaultValue`） |
| 其他 | `toString()` 文本 |

---

## 14. 安全（图片下载）

图片 URL 来自被导出的数据，属**不可信输入**。下载器已内建：仅允许 `http`/`https`、单图 ≤ 64 MB。

数据来源不可信时，开启 SSRF 防护（默认关闭，以保留「内网取图」的合法用法）：

```java
ImageDownloadPolicy.setBlockPrivateNetworks(true);   // 应用启动时设置一次（进程级）
```
开启后拦截解析到环回/链路本地/站点本地/私网/组播地址的 URL，并停止跟随重定向（防 30x 跳内网绕过）。

---

## 15. 枚举与常量

- `ExcelType`：`XLSX` / `XLS`。
- `ExcelCreator` 常量：`XLSX`、`XLS`、`MOVE_AND_RESIZE`(0)、`DONT_MOVE_AND_RESIZE`(3)。
- `AnchorType`：`MOVE_AND_RESIZE`(0)、`MOVE_DONT_RESIZE`(2)、`DONT_MOVE_AND_RESIZE`(3)。
- `CellStyleEnum`：内置单元格样式枚举（用于 DIY 行/列样式）。

---

## 16. 限制与注意事项

- **导入随机访问/图片** 用 `ExcelImportor`/注解导入（DOM）；**超大文件只读** 用 `StreamingExcelReader`（SAX，无图片/随机访问）。
- `autoSizeColumns` 为**抽样估算**（最多采样 1000 行 + 字符宽启发式，CJK 记 2 宽），非像素级；大数据流式模式下仅能看到 SXSSF 窗口内的行，精确控制请显式 `columnWidth`。公式列按其**缓存计算结果**估宽（非公式字符串），故新建未求值的公式列估宽偏窄。
- 流式 `readAsBeans` 按 `@ExcelColumn` **位置**或**表头名**绑定，扩展名缺失的图片 URL 退化为 JPEG。
- 构建/Javadoc 对**含非 ASCII 字符的工程路径**敏感（已在 pom 用 UTF-8 兜底，仍建议纯英文路径）。
- 列字段值通过 getter 反射读取（已用 LambdaMetafactory 缓存加速），行对象需提供对应 getter；导入目标类需对应 setter。
- `@ExcelColumnParent` 仅支持导出（两行表头）；非分组列纵向合并跨两行（序号列跨度>1 时留空；与相邻同名列冲突时该列不纵向合并）；同组子列 index 必须相邻，**不相邻会抛 `ExcelAnnotationException`**。
- `@ExcelInfoChild` 多层嵌套时，每个叶子列按 `父.子.孙` 路径取值/重建；类型自引用成环会被检测并截断该分支。
