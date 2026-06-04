# poi-annotation-excel

A lightweight, annotation-driven Excel library built on Apache POI 5.x with Java 17.

Designed for production use cases where EasyExcel is too opinionated and raw POI is too verbose.
The key differentiator is **multi-level cascade dropdown validation** — a feature absent from most open-source Excel libraries.

---

## Features

| Feature | Description |
|---|---|
| Annotation export & import | `@ExcelColumn`, `@ExcelInfo`, `@ExcelData`, `@ExcelRow` — same model drives both directions |
| Cascade dropdowns | Arbitrary-depth linked dropdowns (e.g. 大类 → 小类 → 设备类型) |
| Inline translation | `@ExcelColumn(translate = {"0:否","1:是"})` — no method needed |
| Path extraction | `@ExcelColumn(sourcePath = "device.type")` for nested objects |
| Template fill | `ExcelTemplateFiller` — `${placeholder}` substitution in templates |
| Row-by-row import | `ExcelReadListener` — per-row callback (DOM-backed) |
| Low-memory streaming import | `StreamingExcelReader` — true SAX streaming for huge `.xlsx`; `read` / `readAsMaps` / `readAsBeans` |
| Image export | Parallel async download, original-format preserving (PNG/JPEG/GIF), disk-staged + STORED ZIP injection — memory-safe for large image sets |
| Auto column width | `autoSizeColumns(true)` — fast sampled estimate (no per-cell font metrics) |
| Custom validation | `ExcelCustomValidate` strategy for row-level business rules |
| CSV support | Transparent CSV read via the same API |
| XLS + XLSX | Both formats supported on export and import |

---

## Requirements

- Java 17+
- Apache POI 5.2.5 (bundled transitively)

---

## Quick Start

### Maven

```xml
<dependency>
    <groupId>io.github.daihu</groupId>
    <artifactId>poi-annotation-excel</artifactId>
    <version>1.0.0</version>
</dependency>
```

### One-liner export — `ExcelUtil` (recommended)

```java
@ExcelInfo(sheetName = "设备列表")
public class DeviceExportModel {

    @ExcelTitle private String title;

    @ExcelData private List<Device> data;

    @ExcelColumn(columnName = "设备名称", index = 1)
    private String devName;

    @ExcelColumn(columnName = "设备状态", index = 2, translate = {"0:离线", "1:在线"})
    private Integer status;

    @ExcelColumn(columnName = "报警级别", index = 3)
    private String alarmLevel;
}

// Single line — model carries data and config
ExcelUtil.export(response.getOutputStream(), "设备列表.xlsx",
        new DeviceExportModel().setTitle("月度报表").setData(deviceList));
```

### Builder export — for dynamic columns (no model class needed)

```java
ExcelUtil.export(response.getOutputStream(), "report.xlsx",
        ExcelCreatorBuilder.create("Sheet1")
                .title("Sales Report")
                .data(salesList)
                .columns("Name:name", "Amount:amount", "Date:date")
                .autoSizeColumns(true));
```

### Cascade dropdown

```java
CascadeValidateModel bigType = CascadeValidateModelBuilder.builder("devBigType")
    .addItem("模拟量设备",
        CascadeValidateModelBuilder.builder("devSmallType")
            .addItem("电流", CascadeValidateModelBuilder.builder("pointTypeCode").addItem("CT1").addItem("CT2"))
            .addItem("电压"))
    .addItem("数字量设备")
    .build();
```

### Batch add from list

```java
builder.addItems(enumDTOs, EnumDTO::getName);  // instead of manual loop
```

### Null-safe children

```java
builder.addItem(name, hasChildren ? childBuilder : null);  // null child = no descendants, no NPE
```

### Template fill

```java
// Template cells: ${reportDate}, ${list.no}, ${list.name}
ExcelTemplateFiller.of(templateInputStream)
    .fill("reportDate", "2024-06-01")
    .fillListBeans("list", rowBeans)     // POJOs work too; or .fillList(...) for List<Map>
    .writeTo(response.getOutputStream());
```

### Import — annotation-driven (recommended)

Use the same `@ExcelInfo` model for both export and import — columns are derived from
`@ExcelColumn` annotations automatically.

```java
@ExcelInfo(sheetName = "设备列表")
public class DeviceImportModel {
    @ExcelColumn(columnName = "设备名称", index = 1) private String devName;
    @ExcelColumn(columnName = "设备状态", index = 2) private Integer status;
    // getters / setters / no-arg constructor required
}

List<DeviceImportModel> devices = ExcelUtil.importExcel(inputStream, DeviceImportModel.class);
```

### Import — list result with manual columns

```java
List<User> users = ExcelUtil.importExcel(inputStream, User.class,
        ExcelModel.of("name"), ExcelModel.of("age"), ExcelModel.of("dept"));

// Skip leading metadata rows (data starts at row 2):
List<User> users = ExcelUtil.importExcel(inputStream, /*sheet=*/0, /*startRow=*/2,
        User.class, ExcelModel.of("name"), ExcelModel.of("age"));

// No POJO needed — get raw maps:
List<Map<String, Object>> rows = ExcelUtil.importExcelToMap(inputStream,
        ExcelModel.of("col1"), ExcelModel.of("col2"));
```

### Streaming import with listener

```java
// Annotation-driven: columns derived from DeviceImportModel.class
ExcelUtil.importExcel(inputStream, DeviceImportModel.class, new ExcelReadListener() {
    @Override public void onRow(Map<String, Object> row, int rowIndex) {
        service.save(convert(row));      // process row immediately — no memory accumulation
    }
    @Override public void onError(String message, int rowIndex) {
        log.warn("Row {} skipped: {}", rowIndex, message);
    }
    @Override public void onFinish(int totalRows) {
        log.info("Import complete, {} rows processed", totalRows);
    }
});

// Manual columns
ExcelUtil.importExcel(inputStream, listener,
        ExcelModel.of("name"), ExcelModel.of("amount"));
```

Supported import target types include `String`, all primitives and their boxed forms,
`BigDecimal`, `BigInteger`, `java.util.Date`, `java.sql.Date`, `java.sql.Timestamp`,
`java.time.LocalDate`, and `java.time.LocalDateTime`. Numeric values tolerate thousands
separators (`"1,234,567"`); date strings tolerate nine common patterns including
ISO-8601, `yyyy-MM-dd HH:mm:ss`, `yyyy/MM/dd`, and `yyyyMMdd`.

### Low-memory streaming import — `StreamingExcelReader` (SAX)

The annotation/listener import above loads the whole workbook into memory (POI DOM). For
very large `.xlsx` files, `StreamingExcelReader` (in `importor/`) reads rows one at a time
via POI's event API, so memory stays bounded regardless of row count.

```java
// 1) Lowest level — handle each row as it streams (nothing accumulates):
StreamingExcelReader.read(inputStream, /*sheet=*/0, (rowIndex, cells) -> {
    service.save(cells);   // cells: List<String> in column order
});

// 2) Header-keyed maps (first row is the header):
List<Map<String, String>> rows = StreamingExcelReader.readAsMaps(inputStream, 0);

// 3) Mapped to an annotated model (columns bind by @ExcelColumn index, first row = header):
List<DeviceImportModel> devices =
        StreamingExcelReader.readAsBeans(inputStream, 0, DeviceImportModel.class);

// 3b) Bind by header name instead of position — column order is irrelevant and
//     unknown columns are ignored (matched against @ExcelColumn.columnName):
List<DeviceImportModel> devices =
        StreamingExcelReader.readAsBeans(inputStream, 0, DeviceImportModel.class, /*matchByHeaderName=*/true);

// Skip leading title/metadata rows — header is at row index 2 (zero-based):
List<DeviceImportModel> devices =
        StreamingExcelReader.readAsBeans(inputStream, 0, DeviceImportModel.class, true, /*headerRowIndex=*/2);

// read / readAsMaps accept the same skip: read from row 5 onward, or header at row 2:
StreamingExcelReader.read(inputStream, 0, /*startRowInclusive=*/5, (i, cells) -> process(cells));
List<Map<String, String>> rows = StreamingExcelReader.readAsMaps(inputStream, 0, /*headerRowIndex=*/2);
```

`readAsBeans` binds columns **by position** (ordered by `@ExcelColumn.index()`) by default, or
**by header name** when the `matchByHeaderName` flag is set; it applies
`@ExcelColumn.translate()` in reverse (display text → stored value), and converts to
`String`, the numeric/boxed types, `BigDecimal`, `boolean`, and the date/time types
(`LocalDate` / `LocalDateTime` / `LocalTime` / `java.util.Date`). Excel serial-date cells
and common text date patterns are both handled.

> Scope: `.xlsx` only, tabular cells only (no images, merged-region expansion, or random
> access); formula cells yield their last cached value. For those features use
> `ExcelImportor`.

A disabled benchmark (`ImportBenchmarkTest`) compares streaming vs DOM import; run it with:

```bash
mvn -o test -Dtest=ImportBenchmarkTest -Dexcel.benchmark=true -Dexcel.benchmark.rows=100000
```

Indicative result for 100,000 × 6 cells (~2 MB file): DOM read ≈ **557 MB** retained heap,
streaming ≈ **0.2 MB** — and the stream is roughly **2–3× faster**. The DOM workbook holds
the whole sheet in memory; the streaming reader holds one row at a time.

---

## Package layout

The public API surface is at the root of `io.github.dh.poi.excel`; everything else is
organised by responsibility:

```
io.github.dh.poi.excel
├── ExcelCreator, ExcelCreatorBuilder, ExcelUtil, ExcelModel, ExcelType   ← public facade
├── annotation/    @ExcelInfo, @ExcelColumn, @ExcelRow, processor
├── cascade/       cascade-validation builder and wrappers
├── core/          workbook strategies, CSV, CellValueSetter, ValueExtractor SPIs
├── exception/     framework exception hierarchy
├── export/        exporter and uploader strategies
├── formula/       Excel formula builders
├── importor/      ExcelImportor + ExcelReadListener
├── model/         ExcelRowData, ComplexExcelModel, DiyRowContextCellModel, RowDataMapper
├── picture/       PictureHandler, AnchorType, PictureConfig
├── render/        cell-resolver chain, ExcelTranslateHandler
├── style/         CellStyleManager, CellStyleEnum
├── template/      ExcelTemplateFiller
└── validation/    DataValidator, ExcelCustomValidate
```

---

## vs. EasyExcel / Hutool ExcelUtil

| | poi-annotation-excel | EasyExcel | Hutool ExcelUtil |
|---|---|---|---|
| Cascade dropdowns | ✅ arbitrary depth | ❌ | ❌ |
| Template fill | ✅ | ✅ | partial |
| Low-memory streaming read | ✅ SAX (`StreamingExcelReader`) | ✅ | ❌ |
| Annotation-driven export | ✅ | ✅ | ✅ |
| Image export | ✅ parallel, format-preserving, memory-safe | limited | ❌ |
| Inline translate mapping | ✅ | ❌ | ❌ |
| Dependency weight | medium | heavy | light |

The cascade dropdown feature is the primary reason to choose this library over alternatives.
EasyExcel has no equivalent; building it manually with raw POI takes hundreds of lines of code.

---

## Building

```bash
mvn clean install
```

Tests require no external services and run fully in-memory.

---

## License

Apache 2.0
