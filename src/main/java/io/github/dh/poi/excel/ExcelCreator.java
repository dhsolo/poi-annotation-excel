/*
 * Copyright 2026 the poi-annotation-excel authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package io.github.dh.poi.excel;

import io.github.dh.poi.excel.annotation.AnnotationProcessor;
import io.github.dh.poi.excel.annotation.DefaultAnnotationProcessor;
import io.github.dh.poi.excel.annotation.ExcelAnnotationProperty;
import io.github.dh.poi.excel.annotation.ExcelInfo;
import io.github.dh.poi.excel.core.BigDataWorkbookStrategy;
import io.github.dh.poi.excel.core.CellValueSetter;
import io.github.dh.poi.excel.core.ValueExtractor;
import io.github.dh.poi.excel.core.WorkbookStrategy;
import io.github.dh.poi.excel.core.XlsWorkbookStrategy;
import io.github.dh.poi.excel.core.XlsxWorkbookStrategy;
import io.github.dh.poi.excel.model.DiyRowContextCellModel;
import io.github.dh.poi.excel.model.ExcelRowData;
import io.github.dh.poi.excel.style.CellStyleEnum;
import io.github.dh.poi.excel.export.DefaultExcelExporter;
import io.github.dh.poi.excel.export.ExcelExporter;
import io.github.dh.poi.excel.export.ExcelUploader;
import io.github.dh.poi.excel.picture.DefaultPictureHandler;
import io.github.dh.poi.excel.picture.PictureHandler;
import io.github.dh.poi.excel.render.CellResolveContext;
import io.github.dh.poi.excel.render.CellValueResolver;
import io.github.dh.poi.excel.render.HandlerCellResolver;
import io.github.dh.poi.excel.render.PictureCellResolver;
import io.github.dh.poi.excel.render.PlainCellResolver;
import io.github.dh.poi.excel.render.TranslateCellResolver;
import io.github.dh.poi.excel.style.CellStyleManager;
import io.github.dh.poi.excel.style.DefaultCellStyleManager;
import io.github.dh.poi.excel.validation.DataValidator;
import io.github.dh.poi.excel.validation.DefaultDataValidator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import io.github.dh.common.Reflect;

import java.io.*;
import java.lang.invoke.CallSite;
import java.lang.invoke.LambdaMetafactory;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.io.Closeable;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * Main entry point for Excel generation.
 *
 * <p>Follows Spring's {@code JdbcTemplate} pattern — this class acts as the primary facade
 * (analogous to {@code ApplicationContext}), assembling and coordinating all sub-modules.
 * Sub-modules depend only on interfaces ({@link CellValueSetter}, {@link ValueExtractor}),
 * never on each other directly.
 *
 * <p>Typical usage:
 * <pre>
 * ExcelCreator creator = ExcelCreatorBuilder.create("Sheet1")
 *     .title("Report")
 *     .header(new String[]{"Name", "Value"})
 *     .data(dataList)
 *     .build();
 * creator.createExcel();
 * creator.export(outputStream, "report.xlsx");
 * </pre>
 *
 * @author dh
 * @since 1.0
 * @see ExcelCreatorBuilder
 * @see CellValueSetter
 * @see ValueExtractor
 */
@SuppressWarnings({"rawtypes", "unused", "unchecked"})
public class ExcelCreator implements CellValueSetter, ValueExtractor, Closeable {

    public static final String XLSX = "xlsx";
    public static final String XLX = "xls";
    public static final int MOVE_AND_RESIZE = 0;
    public static final int MOVE_DONT_RESIZE = 2;
    public static final int DONT_MOVE_AND_RESIZE = 3;

    private static final Logger logger = LoggerFactory.getLogger(ExcelCreator.class);
    private static final int DEFAULT_ROW_HEIGHT = 1000;
    private static final int DEFAULT_COLUMN_WIDTH = 20;
    private static final int DEFAULT_IMAGE_READ_TIME_OUT = 2000;
    private static final AtomicInteger INSTANCE_COUNT = new AtomicInteger(0);
    private static volatile ThreadPoolExecutor executor;
    /**
     * Per-class, per-field getter cache. Each getter is compiled once to a
     * {@link Function} — via {@link LambdaMetafactory} for accessible public getters (near
     * direct-call speed), or a reflective fallback — so repeated cell reads avoid both the
     * hierarchy-walking method lookup and the slower reflective invoke.
     */
    private static final Map<Class<?>, Map<String, Function<Object, Object>>> GETTER_CACHE = new ConcurrentHashMap<>();
    /** Sentinel cached when a class has no matching getter, so negative lookups are not repeated. */
    private static final Function<Object, Object> NO_GETTER = o -> null;

    private static int configuredDownloadThreads = -1;

    /**
     * Configures the number of image download threads. Takes effect globally; call before creating an ExcelCreator.
     * Default: max(cpu cores * 4, 16), capped at 64.
     * <pre>
     * Recommended values:
     *   - Few images (&lt; 50) or low-latency internal network: 16~32
     *   - Many images (&gt; 200) or high-latency public network: 32~64
     *   - Strict server rate limiting: reduce to 8~16
     * </pre>
     */
    public static void setDownloadThreadCount(int count) {
        if (count > 0) configuredDownloadThreads = count;
    }

    static int calculateOptimalThreads() {
        if (configuredDownloadThreads > 0) return configuredDownloadThreads;
        int cpu = Runtime.getRuntime().availableProcessors();
        return Math.max(16, Math.min(cpu * 4, 64));
    }

    private static synchronized ThreadPoolExecutor getOrCreateExecutor() {
        if (executor == null || executor.isShutdown()) {
            int threads = calculateOptimalThreads();
            executor = new ThreadPoolExecutor(threads, threads,
                    1, TimeUnit.MINUTES, new ArrayBlockingQueue<>(threads * 16),
                    r -> new Thread(r, "picture-download-" + java.util.UUID.randomUUID()),
                    new ThreadPoolExecutor.CallerRunsPolicy());
            executor.allowCoreThreadTimeOut(true);
            logger.debug("picture download pool: threads={}, queue={}", threads, threads * 16);
        }
        return executor;
    }

    // ===== Strategy chain: cell rendering =====
    private final List<CellValueResolver> resolvers = Arrays.asList(
            new HandlerCellResolver(),
            new PictureCellResolver(),
            new TranslateCellResolver(),
            new PlainCellResolver()
    );

    /**
     * Per-column resolver cache. The resolver for a column depends only on its
     * {@link ExcelModel}, so it is selected once per column instead of being re-matched
     * (looping {@code resolvers} + calling {@code supports()}) for every cell.
     */
    private final Map<ExcelModel, CellValueResolver> resolverCache = new IdentityHashMap<>();

    // ===== Sub-modules (beans assembled by Spring ApplicationContext) =====
    private CellStyleManager styleManager;
    private PictureHandler pictureHandler;
    private DataValidator dataValidator;
    private ExcelExporter exporter;
    private ExcelCreatePipeline pipeline;

    // ===== Core state =====
    private String currentExcelType;
    private Workbook book;
    private Sheet sheet;
    private Row row;
    private Cell cell;
    private Integer rowHeight;
    private Drawing drawing;
    private int rowNum;
    private boolean excelCreated = false;
    private boolean isBigData = false;

    // ===== Configuration =====
    private Map<String, String[]> headerParentMap = new HashMap<>();
    private String noneCellDefaultValue;
    private int currentListNum = 1;
    private Sheet hiddenSheetListBox;
    private String imagesSeparator = ",";
    private boolean needOrderNum = false;
    private int orderColumnSpan = 1;
    private Object object;
    private String[] header;
    /** Parent-header groups from @ExcelColumnParent (data-column space); null when none. */
    private List<ExcelAnnotationProperty.ParentHeader> parentHeaders;
    private String title;
    private String sheetName;
    private boolean isChildComplex = false;
    private boolean useRootHeaderLength = false;
    private boolean usePrevHeaderLength = false;
    private Integer complexHeaderLength;
    private int titleRowHeight = DEFAULT_ROW_HEIGHT;
    private int headerRowHeight = DEFAULT_ROW_HEIGHT;
    private int imageReadTimeOut;
    private int pictureType = 0;

    // ===== Mapping and merging =====
    private Map<Integer, ExcelModel> columnMappingInfo = new LinkedHashMap<>();
    private Map<String, ExcelModel> columnNameModelMappingInfo = new HashMap<>();
    private Map<Integer, String> columnMergeInfo = new HashMap<>();
    private Map<Integer, List<DiyRowContextCellModel>> diyRowContextCellModelMap = new HashMap<>();
    private Map<Integer, Integer> diyRowContextRowHeightMap = new HashMap<>();

    // ===== Title and custom rows =====
    private CellRangeAddress tileCellRangeAddress;
    private List<CellRangeAddress> diyRowContextCellRangeAddress = new ArrayList<>();

    // ===== Child sheets =====
    private LinkedList<ExcelCreator> child = new LinkedList<>();
    private List<ExcelCreator> complexExcelCreatorList;

    // ===== Validator shared state =====
    private Set<String> existNamaManager = new HashSet<>();
    private final AtomicInteger atomicInteger = new AtomicInteger(0);
    private int currentDiyContextRow;

    // ===== Column widths =====
    private boolean isSettingColumnWidth = false;
    private Map<Integer, Integer> columnWidthMap = new HashMap<>();
    private boolean autoSizeColumns = false;

    // ==================== Constructors ====================

    /**
     * Creates an XLSX {@code ExcelCreator} for the given sheet name.
     *
     * @param sheetName the name of the first sheet; if blank, POI assigns a default name
     */
    public ExcelCreator(String sheetName) {
        this(XLSX, sheetName, false);
    }

    /**
     * Creates an {@code ExcelCreator} from an annotation-driven model object.
     * <p>The object must be annotated with {@code @ExcelInfo}; data, title, headers, and
     * column definitions are all resolved from the annotation metadata automatically.
     *
     * @param excelInfo an instance of a class annotated with {@code @ExcelInfo}
     */
    public ExcelCreator(Object excelInfo) {
        DefaultAnnotationProcessor handler = new DefaultAnnotationProcessor(excelInfo);
        ExcelAnnotationProperty prop = handler.getExcelAnnotationProperty();
        List<AnnotationProcessor> complexModels = prop.getComplexExcelModels();
        if (complexModels != null && complexModels.size() > 0) {
            this.complexExcelCreatorList = complexModels.stream()
                    .map(ExcelCreator::new).collect(Collectors.toList());
        }
        initFromAnnotationProperty(prop, prop.getExcelInfo());
        init(currentExcelType, prop.getExcelInfo().isBigData());
        // Register custom rows declared by @ExcelRow (must be after init(), workbook is ready)
        List<ExcelAnnotationProperty.DiyRowConfig> diyRows = prop.getDiyRows();
        if (diyRows != null) {
            for (var diyRow : diyRows) {
                addDiyRowContext(diyRow.text(), diyRow.merge());
            }
        }
    }

    private ExcelCreator(AnnotationProcessor handler) {
        initFromAnnotationProperty(handler.getExcelAnnotationProperty(),
                handler.getExcelAnnotationProperty().getExcelInfo());
    }

    /**
     * Creates an {@code ExcelCreator} with an explicit file format and big-data flag.
     *
     * @param excelType the file format: {@code "xlsx"} or {@code "xls"}
     * @param sheetName the name of the first sheet; if blank, POI assigns a default name
     * @param bigData   {@code true} to use a streaming (SXSSFWorkbook) strategy for large datasets
     */
    public ExcelCreator(String excelType, String sheetName, boolean bigData) {
        this.sheetName = sheetName;
        init(excelType, bigData);
    }

    /**
     * Creates an {@code ExcelCreator} with an explicit file format, defaulting to non-streaming mode.
     *
     * @param excelType the file format: {@code "xlsx"} or {@code "xls"}
     * @param sheetName the name of the first sheet; if blank, POI assigns a default name
     */
    public ExcelCreator(String excelType, String sheetName) {
        this(excelType, sheetName, false);
    }

    /**
     * Creates an {@code ExcelCreator} that wraps an externally created {@link Workbook}.
     * <p>Useful for attaching this creator to a workbook that already has sheets.
     *
     * @param book an existing POI {@link Workbook} instance
     */
    public ExcelCreator(Workbook book) {
        this.book = book;
    }

    // ==================== CellValueSetter interface implementation ====================

    @Override
    public void setCellValue(Cell cell, Object value) {
        if (cell == null) return;
        if (value == null) { cell.setCellValue(""); return; }
        try {
            // Write numbers and booleans with their native cell type so Excel can sort/aggregate
            // them (rather than storing every value as text). Dates are handled by the
            // date-aware overload below (they need a date-formatted style to display correctly).
            if (value instanceof Number number) {
                cell.setCellValue(number.doubleValue());
            } else if (value instanceof Boolean bool) {
                cell.setCellValue(bool);
            } else {
                cell.setCellValue(value.toString());
            }
        } catch (Exception e) {
            logger.error("Error setting cell value: {}", value, e);
            cell.setCellValue("ERROR");
        }
    }

    /** Cache of date cell styles keyed by format pattern (bounded by the few distinct patterns used). */
    private final Map<String, CellStyle> dateStyleCache = new HashMap<>();

    @Override
    public void setCellValue(Cell cell, Object value, String datePattern) {
        if (cell == null) return;
        if (value != null && isTemporal(value)) {
            try {
                writeTemporal(cell, value, datePattern != null ? datePattern : defaultDatePattern(value));
                return;
            } catch (Exception e) {
                logger.error("Error setting date cell value: {}", value, e);
                cell.setCellValue(value.toString());
                return;
            }
        }
        setCellValue(cell, value);
    }

    private static boolean isTemporal(Object value) {
        return value instanceof java.util.Date
                || value instanceof java.util.Calendar
                || value instanceof java.time.LocalDate
                || value instanceof java.time.LocalDateTime;
    }

    private static String defaultDatePattern(Object value) {
        return (value instanceof java.time.LocalDate || value instanceof java.sql.Date)
                ? "yyyy-MM-dd" : "yyyy-MM-dd HH:mm:ss";
    }

    private void writeTemporal(Cell cell, Object value, String pattern) {
        CellStyle style = dateStyleCache.computeIfAbsent(pattern, p -> {
            CellStyle s = book.createCellStyle();
            s.cloneStyleFrom(styleManager.getCellStyle());
            s.setDataFormat(book.getCreationHelper().createDataFormat().getFormat(p));
            return s;
        });
        cell.setCellStyle(style);
        if (value instanceof java.util.Date d) {
            cell.setCellValue(d);
        } else if (value instanceof java.util.Calendar c) {
            cell.setCellValue(c);
        } else if (value instanceof java.time.LocalDate ld) {
            cell.setCellValue(ld);
        } else if (value instanceof java.time.LocalDateTime ldt) {
            cell.setCellValue(ldt);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    @Override
    @SuppressWarnings("unchecked")
    public ExcelRowData createExcelRowData(Object obj, String field, int colIndex, int rowIndex) {
        return new ExcelRowData<Object>() {
            @Override public Row getRow() { return row; }
            @Override public Cell getCell() { return cell; }
            @Override public Object getRowData() { return obj; }
            @Override public <U> U getRowData(Class<U> uClass) { return uClass.cast(obj); }
            @Override public Object currentValue() { return getValue(field, obj); }
            @Override public int currentCellNum() { return colIndex; }
            @Override public int currentRowNum() { return rowNum + rowIndex; }
            @Override public <U> U getOriginalCellData(Class<U> uClass) { return null; }
            @Override public Map<String, Object> getOriginalCellData() { return null; }
        };
    }

    // ==================== ValueExtractor interface implementation ====================

    @Override
    public Object getValue(String field, Object data) {
        if (data instanceof Map<?, ?> map) {
            if (map.containsKey(field)) return map.get(field);
            for (var key : map.keySet()) {
                if (key != null && field.equals(key.toString())) return map.get(key);
            }
            return null;
        }
        // Resolve the getter through a per-class cache so each cell read is a map lookup plus a
        // compiled getter call instead of a hierarchy-walking lookup and reflective invoke.
        Function<Object, Object> getter = GETTER_CACHE
                .computeIfAbsent(data.getClass(), k -> new ConcurrentHashMap<>())
                .computeIfAbsent(field, f -> buildGetter(data.getClass(), f));
        if (getter == NO_GETTER) return null;
        try {
            return getter.apply(data);
        } catch (Exception e) {
            logger.error("Reflection error while getting data value", e);
            return "";
        }
    }

    /**
     * Compiles a getter for {@code field} on {@code cls} into a {@link Function}. Accessible
     * public getters are turned into a {@link LambdaMetafactory} lambda (near direct-call
     * speed); everything else falls back to a reflective invoke. Returns {@link #NO_GETTER}
     * when no matching getter exists.
     */
    private static Function<Object, Object> buildGetter(Class<?> cls, String field) {
        String getMethod = "get" + Character.toUpperCase(field.charAt(0)) + field.substring(1);
        Method m = Reflect.findMethod(cls, getMethod);
        if (m == null) return NO_GETTER;

        if (Modifier.isPublic(m.getModifiers()) && Modifier.isPublic(cls.getModifiers())) {
            try {
                MethodHandles.Lookup lookup = MethodHandles.lookup();
                MethodHandle mh = lookup.unreflect(m);
                CallSite site = LambdaMetafactory.metafactory(
                        lookup, "apply",
                        MethodType.methodType(Function.class),
                        MethodType.methodType(Object.class, Object.class),
                        mh, mh.type());
                @SuppressWarnings("unchecked")
                Function<Object, Object> fn = (Function<Object, Object>) site.getTarget().invokeExact();
                return fn;
            } catch (Throwable ignored) {
                // fall through to reflective getter (e.g. inaccessible across modules)
            }
        }

        Method reflective = m; // accessibility already forced by findMethod
        return obj -> {
            try {
                return reflective.invoke(obj);
            } catch (IllegalAccessException | InvocationTargetException e) {
                throw new RuntimeException(e);
            }
        };
    }

    // ==================== Initialization ====================

    private void init(String excelType, boolean bigData) {
        checkExcelType(excelType);
        createWorkBook(resolveWorkbookStrategy(excelType, bigData));
    }

    private WorkbookStrategy resolveWorkbookStrategy(String excelType, boolean bigData) {
        if (bigData) return new BigDataWorkbookStrategy();
        if (XLSX.equals(excelType)) return new XlsxWorkbookStrategy();
        return new XlsWorkbookStrategy();
    }

    private void initFromAnnotationProperty(ExcelAnnotationProperty prop, ExcelInfo info) {
        setTitle(prop.getTitle());
        setHeader(prop.getHeader());
        this.parentHeaders = prop.getParentHeaders();
        setObject(prop.getExcelData());
        setColumnMergeInfo(prop.getMergeInfo());
        Map<Integer, ExcelModel> mapping = new HashMap<>();
        List<ExcelModel> models = prop.getExcelModels();
        if (models != null && models.size() > 0) {
            models.forEach(m -> {
                mapping.put(mapping.size(), m);
                m.setRealIndex(info.needOrder()
                        ? mapping.size() + 1 + (orderColumnSpan - 1)
                        : mapping.size());
            });
        }
        setColumnMappingInfo(mapping);
        Map<Integer, Integer> widthInfo = prop.getColumnWidthInfo();
        if (widthInfo != null && !widthInfo.isEmpty()) {
            widthInfo.forEach((col, charWidth) -> columnWidthMap.put(col, charWidth * 255));
        }
        this.noneCellDefaultValue = info.noneCellDefaultValue();
        this.sheetName = info.sheetName();
        this.currentExcelType = info.excelType();
        this.setNeedOrderNum(info.needOrder());
        this.orderColumnSpan = info.orderColumnSpan();
        this.pictureType = info.pictureInnerType();
        this.imageReadTimeOut = info.imageReadTimeOut();
        setTitleRowHeight(info.titleHeight());
        setHeaderRowHeight(info.headerHeight());
        setImagesSeparator(Reflect.hasText(info.imageSeparator()) ? info.imageSeparator() : ",");
    }

    private void initHelpers() {
        styleManager = new DefaultCellStyleManager(book);
        styleManager.initDefaultStyles();

        INSTANCE_COUNT.incrementAndGet();
        pictureHandler = new DefaultPictureHandler(book, currentExcelType, sheet, drawing,
                imagesSeparator, getOrCreateExecutor(), this);
        pictureHandler.setPictureType(pictureType);
        pictureHandler.setImageReadTimeOut(imageReadTimeOut);

        dataValidator = new DefaultDataValidator(book, sheet, currentExcelType, hiddenSheetListBox,
                isBigData, existNamaManager, atomicInteger, columnNameModelMappingInfo);
        dataValidator.setCurrentListNum(currentListNum);

        exporter = new DefaultExcelExporter(book, currentExcelType);
        pipeline = new ExcelCreatePipeline(this);
    }

    // ==================== Core: Excel creation ====================

    /**
     * Executes the Excel generation pipeline: writes the title row, custom rows, header row,
     * data rows, cell merges, child sheets, and picture embedding in the correct order.
     * <p>This method must be called before any export or upload operation.
     * It is safe to call only once per instance; subsequent calls are a no-op when a pipeline
     * has already been executed.
     *
     * @return {@code this} for fluent chaining (e.g. {@code creator.createExcel().export(out, name)})
     */
    public ExcelCreator createExcel() {
        if (excelCreated) return this;
        if (pipeline != null) {
            pipeline.execute();
        } else {
            legacyCreateExcel();
        }
        return this;
    }

    private void legacyCreateExcel() {
        logger.debug("create excel start");
        long start = System.currentTimeMillis();
        if (!isChildComplex && pictureHandler != null) pictureHandler.createTempleFileDir();
        if (pictureHandler != null) preparePictureData();
        writeHeaderAndTitle(true);
        if (dataValidator != null) {
            dataValidator.checkListBox(columnMappingInfo, needOrderNum, rowNum);
            dataValidator.setRowNum(rowNum);
            currentListNum = dataValidator.getCurrentListNum();
        }
        if (pictureHandler != null) pictureHandler.downLoadPicture();
        populateData();
        mergeCells();
        createChildSheets();
        if (!isChildComplex && pictureHandler != null) {
            pictureHandler.decompressionPictureDirAndCompression();
            if (exporter != null) {
                exporter.setZip(pictureHandler.isZip());
                exporter.setTempWorkFile(pictureHandler.getTempWorkFile());
            }
        }
        long duration = System.currentTimeMillis() - start;
        double time = new BigDecimal(duration).divide(new BigDecimal(1000), 8, RoundingMode.DOWN).doubleValue();
        logger.debug("create excel finished , spend time:{}", time);
        excelCreated = true;
    }

    // ==================== Pictures ====================

    void preparePictureData() {
        List<?> dataList = resolveDataList();
        pictureHandler.checkPictureMaxSize(columnMappingInfo, dataList, header, needOrderNum, orderColumnSpan);
        String[] expandedHeader = pictureHandler.expandHeaderForPictures(header);
        if (expandedHeader != header) header = expandedHeader;
    }

    /**
     * Selects the {@link CellValueResolver} for a column, caching the result so the resolver
     * list is matched only once per column rather than per cell.
     *
     * @param excelModel the column metadata
     * @return the first resolver that supports the column, or {@code null} if none does
     */
    private CellValueResolver resolverFor(ExcelModel excelModel) {
        CellValueResolver cached = resolverCache.get(excelModel);
        if (cached != null) return cached;
        for (CellValueResolver resolver : resolvers) {
            if (resolver.supports(excelModel)) {
                resolverCache.put(excelModel, resolver);
                return resolver;
            }
        }
        return null;
    }

    // ==================== Data population (strategy pattern) ====================

    void populateData() {
        List<?> data = resolveDataList();
        CellStyle dataCellStyle = styleManager.getCellStyle();

        for (int i = 0; i < data.size(); i++) {
            Object obj = data.get(i);
            row = sheet.getRow(rowNum + i);
            if (row == null) row = sheet.createRow(rowNum + i);
            row.setHeight(rowHeight != null ? rowHeight.shortValue() : (short) DEFAULT_ROW_HEIGHT);

            int length = header.length;
            if (needOrderNum) length += 1 + (orderColumnSpan - 1);
            int max = 0;

            for (int j = 0; j < length; j++) {
                cell = row.getCell(j);
                if (cell == null) cell = row.createCell(j);
                cell.setCellStyle(dataCellStyle);

                if (needOrderNum && j < orderColumnSpan) {
                    cell.setCellValue(i + 1);
                } else {
                    ExcelModel excelModel;
                    if (needOrderNum) {
                        excelModel = columnMappingInfo.get(j - max - 1 - (orderColumnSpan - 1));
                    } else {
                        excelModel = columnMappingInfo.get(j - max);
                    }

                    CellResolveContext ctx = new CellResolveContext(
                            this, this, cell, row, sheet, excelModel, obj, j, i,
                            rowNum, noneCellDefaultValue, dataCellStyle, pictureHandler);

                    CellValueResolver resolver = resolverFor(excelModel);
                    if (resolver != null) {
                        max = resolver.resolve(ctx);
                    }
                    j += max;

                    if (needOrderNum && j == orderColumnSpan && orderColumnSpan > 1) {
                        setMergeColumn(rowNum + i, rowNum + i, 0, orderColumnSpan - 1);
                    }
                    if (excelModel.getMergeCellIndex() > 1 && !excelModel.getMergeIndexEnd()) {
                        setMergeColumn(rowNum + i, rowNum + i, j, j + excelModel.getMergeCellIndex() - 1);
                    }
                }
            }
        }
        rowNum += data.size();
        handleComplexCreator();
        logger.debug("Data population complete");
    }

    private void handleComplexCreator() {
        if (complexExcelCreatorList == null || complexExcelCreatorList.isEmpty()) return;
        AtomicReference<ExcelCreator> ex = new AtomicReference<>();
        complexExcelCreatorList.forEach(ec -> {
            ec.currentDiyContextRow = currentDiyContextRow;
            ec.isChildComplex = true;
            ec.sheet = sheet;
            ec.rowNum = rowNum;
            ec.book = book;
            ec.child = null;
            ec.drawing = drawing;
            ec.currentExcelType = currentExcelType;
            ec.hiddenSheetListBox = hiddenSheetListBox;
            ec.currentListNum = currentListNum;
            ec.existNamaManager = existNamaManager;
            ec.atomicInteger.set(atomicInteger.get());
            if (pictureHandler != null) ec.pictureHandler = pictureHandler;
            ec.defaultCellStyle();
            ec.createExcel();
            if (pictureHandler != null && !pictureHandler.hasPicture() && ec.pictureHandler.hasPicture()) {
                pictureHandler.setHasPicture(true);
            }
            rowNum = ec.rowNum;
            currentDiyContextRow = ec.currentDiyContextRow;
            ex.set(ec);
        });
    }

    // ==================== Header and title ====================

    void writeHeaderAndTitle(boolean needHandle) {
        CellStyle titleStyle = styleManager.getTitleCellStyle();
        CellStyle headerStyle = styleManager.getHeaderCellStyle();
        diyRowContextCellRangeAddress.clear();

        if (title != null && title.trim().length() > 0) {
            int dsize = 0;
            for (Integer in : diyRowContextCellModelMap.keySet()) {
                List<DiyRowContextCellModel> models = diyRowContextCellModelMap.get(in);
                long count = models.stream().filter(f -> !f.isAfterTitle()).count();
                if (count > 0) {
                    if (needHandle) {
                        Integer height = diyRowContextRowHeightMap.get(in);
                        row = sheet.createRow(rowNum);
                        row.setHeight(height != null ? height.shortValue() : (short) titleRowHeight);
                        setText(dsize, models, getHeaderLength() - 1 - (orderColumnSpan - 1));
                    }
                    rowNum++;
                    dsize++;
                }
            }
            row = sheet.createRow(rowNum);
            row.setHeight((short) titleRowHeight);
            if (needHandle) {
                int hl = getHeaderLength();
                for (int i = 0; i < hl; i++) {
                    cell = row.createCell(i);
                    cell.setCellStyle(titleStyle);
                    cell.setCellValue(title);
                }
                tileCellRangeAddress = new CellRangeAddress(rowNum, rowNum, 0, hl - 1 - (orderColumnSpan - 1));
            } else {
                cell = row.createCell(0);
                cell.setCellStyle(titleStyle);
                cell.setCellValue(title);
            }
            rowNum++;
        }

        if (diyRowContextCellModelMap.size() > 0) {
            int dsize = 0;
            for (int i = 0; i < diyRowContextCellModelMap.size(); i++) {
                List<DiyRowContextCellModel> models = diyRowContextCellModelMap.get(i);
                if (models.stream().filter(f -> !f.isAfterTitle()).count() > 0) continue;
                if (needHandle) {
                    Integer height = diyRowContextRowHeightMap.get(i);
                    row = sheet.createRow(rowNum);
                    row.setHeight(height != null ? height.shortValue() : (short) titleRowHeight);
                    setText(dsize, models, getHeaderLength() - 1 - (orderColumnSpan - 1));
                }
                rowNum++;
                dsize++;
            }
        }

        // Parent (grouped) header row from @ExcelColumnParent: rendered just above the column
        // header row, with each parent label merged across its child column range.
        if (parentHeaders != null && !parentHeaders.isEmpty() && needHandle
                && header != null && header.length > 0) {
            int orderOffset = needOrderNum ? orderColumnSpan : 0;
            int plength = header.length + orderOffset;
            int childHeaderRow = rowNum + 1; // the column-header row written by the block below
            Row pr = sheet.createRow(rowNum);
            pr.setHeight((short) headerRowHeight);
            for (int i = 0; i < plength; i++) {
                pr.createCell(i).setCellStyle(headerStyle);
            }

            boolean[] grouped = new boolean[plength];
            for (ExcelAnnotationProperty.ParentHeader ph : parentHeaders) {
                int s = ph.startCol() + orderOffset;
                int e = Math.min(ph.endCol() + orderOffset, plength - 1);
                if (s >= plength) continue;
                pr.getCell(s).setCellValue(ph.label());
                if (e > s) setMergeColumn(rowNum, rowNum, s, e);
                for (int c = s; c <= e; c++) grouped[c] = true;
            }

            // Non-grouped columns span both header rows: put the name in the (top) parent row and
            // vertically merge down so a single label covers both rows. Only when the order column
            // mapping is unambiguous (no order column, or a single order column).
            if (!needOrderNum || orderColumnSpan == 1) {
                for (int i = 0; i < plength; i++) {
                    if (grouped[i]) continue;
                    String name = (needOrderNum && i == 0) ? "序号" : header[i - orderOffset];
                    pr.getCell(i).setCellValue(name);
                    setMergeColumn(rowNum, childHeaderRow, i, i);
                }
            }
            rowNum++;
        }

        if (header != null && header.length > 0 && needHandle) {
            row = sheet.createRow(rowNum);
            row.setHeight((short) headerRowHeight);
            int length = header.length;
            if (needOrderNum) length += 1 + (orderColumnSpan - 1);
            String headerName = null;
            int startColumn = 0, endColumn = 0;
            for (int i = 0; i < length; i++) {
                boolean isMerge = false;
                cell = row.createCell(i);
                cell.setCellStyle(headerStyle);
                Integer cw = columnWidthMap.get(i);
                sheet.setColumnWidth(i, cw != null ? cw : DEFAULT_COLUMN_WIDTH * 255);
                if (needOrderNum && i < orderColumnSpan) {
                    cell.setCellValue("序号");
                } else if (needOrderNum && i == orderColumnSpan && orderColumnSpan > 1) {
                    setMergeColumn(rowNum, rowNum, 0, orderColumnSpan - 1);
                } else {
                    int idx = needOrderNum ? i - 1 : i;
                    if (headerName == null) {
                        headerName = header[idx];
                        startColumn = endColumn = i;
                    } else if (headerName.equals(header[idx])) {
                        endColumn++;
                    } else if (!Objects.equals(headerName, header[idx])) {
                        if (startColumn != endColumn) {
                            mergeCellWidthHandle(startColumn, endColumn);
                            setMergeColumn(rowNum, rowNum, startColumn, endColumn);
                            isMerge = true;
                        }
                        startColumn = endColumn = i;
                        headerName = header[idx];
                    }
                    cell.setCellValue(header[idx]);
                }
                if (i == length - 1 && startColumn != endColumn && !isMerge) {
                    mergeCellWidthHandle(startColumn, endColumn);
                    setMergeColumn(rowNum, rowNum, startColumn, endColumn);
                }
            }
            rowNum++;
        }
    }

    private int getHeaderLength() {
        int hl = (header != null && header.length > 0) ? header.length :
                (complexHeaderLength != null ? complexHeaderLength : 0);
        if (needOrderNum) hl += 1 + (orderColumnSpan - 1);
        return hl;
    }

    private void setText(int dsize, List<DiyRowContextCellModel> models, int size) {
        CellStyle titleStyle = styleManager.getTitleCellStyle();
        CellStyle dataStyle = styleManager.getCellStyle();
        for (var m : models) {
            int sc = m.getStartColumn() == -1 ? 0 : m.getStartColumn();
            int ec = m.getEndColumn() == -1 ? size : m.getEndColumn();
            CellStyle cs = switch (m.getCellStyleEnum()) {
                case customStyle -> m.getCustomCellStyle();
                case titleStyle, defaultStyle -> titleStyle;
                case normalStyle -> dataStyle;
            };
            for (int i = sc; i <= ec; i++) {
                cell = row.createCell(i);
                cell.setCellStyle(cs);
                cell.setCellValue(m.getValue());
            }
            if (sc != ec) diyRowContextCellRangeAddress.add(new CellRangeAddress(rowNum, rowNum, sc, ec));
        }
    }

    private void mergeCellWidthHandle(int startColumn, int endColumn) {
        Integer width = columnWidthMap.get(startColumn);
        int w = width != null ? width : DEFAULT_COLUMN_WIDTH * 255;
        for (int i = startColumn; i <= endColumn; i++) sheet.setColumnWidth(i, w);
    }

    // ==================== Cell merging ====================

    void mergeCells() {
        if (tileCellRangeAddress != null) sheet.addMergedRegion(tileCellRangeAddress);
        if (!diyRowContextCellRangeAddress.isEmpty())
            diyRowContextCellRangeAddress.forEach(f -> sheet.addMergedRegion(f));
        int baseRowNum = 0;
        if (title != null && title.trim().length() != 0) baseRowNum++;
        if (header != null && header.length != 0) baseRowNum++;
        baseRowNum += diyRowContextCellModelMap.size();
        if (columnMergeInfo != null) {
            for (Map.Entry<Integer, String> entry : columnMergeInfo.entrySet()) {
                int indexColumn = entry.getKey();
                int index = 0;
                if (needOrderNum) indexColumn++;
                String field = entry.getValue();
                List dataList = resolveDataList();
                Object currentValue = null;
                for (int i = 0; i < dataList.size(); i++) {
                    Object value = getValue(field, dataList.get(i));
                    if (currentValue == null) {
                        currentValue = value;
                    } else if (!(currentValue + "").equals(value + "")) {
                        // Value changed: merge the [index, i-1] range
                        int firstRow = index + baseRowNum;
                        int lastRow = i - 1 + baseRowNum;
                        if (firstRow != lastRow) {
                            setMergeColumn(firstRow, lastRow, indexColumn, indexColumn);
                        }
                        index = i;
                        currentValue = value;
                    }
                }
                // Handle the trailing group of consecutive equal values (merge range not yet committed when the loop ends)
                if (currentValue != null) {
                    int firstRow = index + baseRowNum;
                    int lastRow = dataList.size() - 1 + baseRowNum;
                    if (firstRow != lastRow) {
                        setMergeColumn(firstRow, lastRow, indexColumn, indexColumn);
                    }
                }
            }
        }
        applyAutoSizeColumns();
    }

    /**
     * Merges a rectangular cell region in the active sheet.
     * <p>All row and column indices are zero-based.
     *
     * @param firstRow the first row of the merge range (inclusive)
     * @param lastRow  the last row of the merge range (inclusive)
     * @param firstCol the first column of the merge range (inclusive)
     * @param lastCol  the last column of the merge range (inclusive)
     */
    public void setMergeColumn(int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    // ==================== Child sheets ====================

    void createChildSheets() {
        if (child == null || child.isEmpty()) return;
        int concurrentRowNum = currentListNum;
        for (ExcelCreator ec : child) {
            ec.currentListNum = concurrentRowNum;
            ec.existNamaManager = existNamaManager;
            ec.atomicInteger.set(atomicInteger.get());
            ec.book = book;
            ec.sheet = ec.sheetName != null && ec.sheetName.length() > 0
                    ? ec.book.createSheet(ec.sheetName) : ec.book.createSheet();
            ec.drawing = ec.sheet.createDrawingPatriarch();
            ec.currentExcelType = currentExcelType;
            ec.hiddenSheetListBox = hiddenSheetListBox;
            ec.defaultCellStyle();
            ec.createExcel();
            concurrentRowNum = ec.currentListNum;
        }
        currentListNum = concurrentRowNum;
    }

    /**
     * Returns the child {@code ExcelCreator} at the given position within a complex (multi-section) layout.
     * <p>Complex creators are declared via {@code @ExcelInfo} on nested model fields.
     *
     * @param index zero-based index of the complex section
     * @return the complex {@code ExcelCreator}, or {@code null} if the index is out of bounds
     *         or no complex sections were declared
     */
    public ExcelCreator getComplexCreator(int index) {
        if (complexExcelCreatorList == null || complexExcelCreatorList.isEmpty()
                || index < 0 || index >= complexExcelCreatorList.size()) return null;
        return complexExcelCreatorList.get(index);
    }

    private List<?> resolveDataList() {
        List<Object> data = new LinkedList<>();
        if (object instanceof Map<?, ?> m) data.add(m);
        else if (object instanceof List<?> l) data.addAll(l);
        return data;
    }

    // ==================== Workbook creation (strategy pattern) ====================

    private void checkExcelType(String excelType) {
        if (excelType == null || excelType.trim().length() == 0)
            throw new IllegalArgumentException("Excel type must not be blank");
        if (!XLSX.equals(excelType) && !XLX.equals(excelType))
            throw new IllegalArgumentException("Excel type must be 'xlsx' or 'xls'");
    }

    private void createWorkBook(WorkbookStrategy strategy) {
        book = strategy.createWorkbook();
        currentExcelType = strategy.getExcelType();
        isBigData = strategy.isBigData();
        sheet = (sheetName != null && sheetName.trim().length() > 0)
                ? book.createSheet(sheetName) : book.createSheet();
        drawing = sheet.createDrawingPatriarch();
        imageReadTimeOut = DEFAULT_IMAGE_READ_TIME_OUT;
        hiddenSheetListBox = book.createSheet("listConstantData");
        book.setSheetHidden(book.getSheetIndex(hiddenSheetListBox), true);
        initHelpers();
        defaultCellStyle();
    }

    /**
     * Re-initialises the workbook with the given format and big-data flag.
     * <p>Intended for advanced use cases where the workbook strategy must be changed after
     * construction (e.g. switching between XLSX and big-data streaming mode at runtime).
     * Calling this method replaces the current workbook and resets all internal state.
     *
     * @param excelType the file format: {@code "xlsx"} or {@code "xls"}
     * @param bigData   {@code true} to use a streaming (SXSSFWorkbook) strategy
     */
    public void createWorkBook(String excelType, boolean bigData) {
        createWorkBook(resolveWorkbookStrategy(excelType, bigData));
    }

    private void defaultCellStyle() {
        // Rebuild the style manager against the CURRENT workbook. This matters when a child
        // creator is stitched into a parent's workbook (multi-sheet): cell styles belong to a
        // specific workbook, so styles created against the child's original book cannot be used
        // in the parent's book ("Style does not belong to the supplied Workbook").
        styleManager = new DefaultCellStyleManager(book);
        styleManager.initDefaultStyles();
        dateStyleCache.clear(); // date styles cached against the previous workbook are now invalid
    }

    // ==================== Export ====================

    /**
     * Returns the underlying POI {@link Workbook} managed by the current exporter.
     *
     * @return the active {@link Workbook} instance
     */
    public Workbook getWorkBook() { return exporter != null ? exporter.getWorkBook() : book; }

    /**
     * Writes the fully-built workbook to the given output stream.
     * <p>{@link #createExcel()} must be called before this method.
     *
     * @param os the output stream to write to (e.g. an HTTP response output stream)
     * @param fn the file name (used by the exporter; may influence the {@code Content-Disposition} header)
     * @throws IOException if an I/O error occurs while writing
     */
    public void export(OutputStream os, String fn) throws IOException { if (exporter != null) exporter.export(os, fn); }

    /**
     * Uploads the fully-built workbook via the supplied {@link ExcelUploader} strategy.
     * <p>{@link #createExcel()} must be called before this method.
     *
     * @param uploader the upload strategy (e.g. S3, MinIO, or a custom file service)
     * @param fn       the target file name (extension recommended, e.g. {@code "report.xlsx"})
     * @param <R>      the upload result type returned by the uploader (e.g. a URL or file ID)
     * @return the result produced by {@link ExcelUploader#upload}, or {@code null} if no exporter is configured
     * @throws IOException if an I/O error occurs during upload
     */
    public <R> R upload(ExcelUploader<R> uploader, String fn) throws IOException { return exporter != null ? exporter.upload(uploader, fn) : null; }

    /**
     * Returns the workbook content as an {@link InputStream}, suitable for streaming uploads (e.g. OSS/S3).
     * <p>{@link #createExcel()} must be called before this method.
     *
     * @return an {@link InputStream} over the serialised workbook bytes
     * @throws RuntimeException if {@link #createExcel()} has not yet been called
     */
    public InputStream getInputStream() {
        if (exporter != null) return exporter.getInputStream(excelCreated);
        throw new RuntimeException("Excel has not been created yet; call createExcel() first");
    }

    /**
     * Saves the workbook to a local file at the specified path.
     * <p>{@link #createExcel()} must be called before this method.
     *
     * @param fp the absolute or relative file path (no extension needed; the exporter appends it)
     */
    public void exportLocal(String fp) { if (exporter != null) exporter.exportLocal(fp); }

    // ==================== Styles ====================

    /**
     * Returns a new {@link CellStyle} copied from the current title style, ready for customisation.
     *
     * @return a copy of the title cell style, or {@code null} if the style manager is not initialised
     */
    public CellStyle copyTitleStyle() { return styleManager != null ? styleManager.copyTitleStyle() : null; }

    /**
     * Returns a new {@link CellStyle} copied from the current data-cell style, ready for customisation.
     *
     * @return a copy of the data cell style, or {@code null} if the style manager is not initialised
     */
    public CellStyle copyCellStyle() { return styleManager != null ? styleManager.copyCellStyle() : null; }

    /**
     * Returns a new {@link CellStyle} copied from the current header style, ready for customisation.
     *
     * @return a copy of the header cell style, or {@code null} if the style manager is not initialised
     */
    public CellStyle copyHeaderStyle() { return styleManager != null ? styleManager.copyHeaderStyle() : null; }

    /**
     * Replaces the default title row cell style with a custom one.
     *
     * @param s the custom {@link CellStyle} to apply to title cells
     */
    public void setTitleCellStyle(CellStyle s) { if (styleManager != null) styleManager.setTitleCellStyle(s); }

    /**
     * Replaces the default data cell style with a custom one.
     *
     * @param s the custom {@link CellStyle} to apply to data cells
     */
    public void setCellStyle(CellStyle s) { if (styleManager != null) styleManager.setCellStyle(s); }

    /**
     * Replaces the default header row cell style with a custom one.
     *
     * @param s the custom {@link CellStyle} to apply to header cells
     */
    public void setHeaderCellStyle(CellStyle s) { if (styleManager != null) styleManager.setHeaderCellStyle(s); }

    /**
     * Creates a new, blank {@link CellStyle} in the underlying workbook.
     * <p>Use this together with {@link #createFont()} to build fully custom styles.
     *
     * @return a new empty {@link CellStyle} registered in the workbook
     */
    public CellStyle createCellStyle() { return getWorkBook().createCellStyle(); }

    /**
     * Creates a new {@link Font} in the underlying workbook for use in custom cell styles.
     *
     * @return a new {@link Font} registered in the workbook
     */
    public Font createFont() { return getWorkBook().createFont(); }

    // ==================== Custom row information ====================

    /**
     * Appends a fully-merged custom row above the header, using the default title row height.
     * <p>The row spans all header columns and uses the title cell style.
     *
     * @param context the text to display in the custom row
     */
    public void addDiyRowContext(String context) { addDiyRowContext(context, titleRowHeight); }

    /**
     * Appends a fully-merged custom row above the header with an explicit row height.
     *
     * @param context the text to display in the custom row
     * @param height  row height in twips (1/20th of a point); e.g. {@code 1000} ≈ 50px
     */
    public void addDiyRowContext(String context, int height) { addDiyRowContext(context, true, height); }

    /**
     * Appends a custom row cell spanning columns {@code sc}–{@code ec} with the default title row height.
     *
     * @param context the text to display
     * @param sc      start column index (inclusive, 0-based; pass {@code -1} to span from the first column)
     * @param ec      end column index (inclusive, 0-based; pass {@code -1} to span to the last header column)
     * @param cse     cell style preset; {@code null} falls back to title style
     * @param cs      fully custom {@link CellStyle} used when {@code cse} is {@link CellStyleEnum#customStyle}
     * @param is      {@code true} to append an additional cell to the same logical row (not yet committed);
     *                {@code false} to commit the row and advance the row counter
     */
    public void addDiyRowContext(String context, int sc, int ec, CellStyleEnum cse, CellStyle cs, boolean is) { addDiyRowContext(context, sc, ec, cse, cs, is, titleRowHeight); }

    /**
     * Appends a custom row cell with explicit placement relative to the title row.
     *
     * @param context    the text to display
     * @param sc         start column index (inclusive, 0-based; {@code -1} = first column)
     * @param ec         end column index (inclusive, 0-based; {@code -1} = last header column)
     * @param cse        cell style preset
     * @param cs         fully custom {@link CellStyle} for {@link CellStyleEnum#customStyle}
     * @param is         {@code true} to keep the row open for more cells; {@code false} to commit
     * @param at         {@code true} to insert this row <em>after</em> the title row;
     *                   {@code false} to insert it <em>before</em> the title row
     */
    public void addDiyRowContext(String context, int sc, int ec, CellStyleEnum cse, CellStyle cs, boolean is, boolean at) { addDiyRowContext(context, sc, ec, cse, cs, is, titleRowHeight, at); }

    /**
     * Appends a custom row cell spanning columns {@code sc}–{@code ec} with an explicit row height.
     *
     * @param context the text to display
     * @param sc      start column index (inclusive, 0-based; {@code -1} = first column)
     * @param ec      end column index (inclusive, 0-based; {@code -1} = last header column)
     * @param cse     cell style preset
     * @param cs      fully custom {@link CellStyle} for {@link CellStyleEnum#customStyle}
     * @param is      {@code true} to keep the row open for more cells; {@code false} to commit
     * @param h       row height in twips
     */
    public void addDiyRowContext(String context, int sc, int ec, CellStyleEnum cse, CellStyle cs, boolean is, int h) { addDiyRowContext(context, sc, ec, cse, cs, is, h, true); }

    /**
     * Core implementation for adding a custom row cell.
     * <p>Appends a {@link DiyRowContextCellModel} to the current logical row.
     * When {@code is} is {@code false} the current row is finalised (row height recorded,
     * row counter incremented) so subsequent calls start a new row.
     *
     * @param context the text to display
     * @param sc      start column (0-based; {@code -1} = first)
     * @param ec      end column (0-based; {@code -1} = last header column)
     * @param cse     cell style preset; may be {@code null}
     * @param cs      custom {@link CellStyle} for {@link CellStyleEnum#customStyle}; may be {@code null}
     * @param is      {@code true} to continue the current row; {@code false} to commit it
     * @param h       row height in twips when committing the row
     * @param at      {@code true} to place the row after the title; {@code false} to place it before
     */
    public void addDiyRowContext(String context, int sc, int ec, CellStyleEnum cse, CellStyle cs, boolean is, int h, boolean at) {
        List<DiyRowContextCellModel> models = diyRowContextCellModelMap.computeIfAbsent(currentDiyContextRow, k -> new ArrayList<>());
        DiyRowContextCellModel m = new DiyRowContextCellModel();
        m.setStartColumn(sc); m.setEndColumn(ec); m.setValue(context); m.setAfterTitle(at);
        if (cse != null) m.setCellStyleEnum(cse);
        if (cs != null) m.setCustomCellStyle(cs);
        models.add(m);
        if (!is) { diyRowContextRowHeightMap.put(currentDiyContextRow, h); currentDiyContextRow++; }
    }

    /**
     * Appends a custom row with optional full-width merging, using the default title row height.
     *
     * @param context    the text to display
     * @param needMerge  {@code true} to merge all columns in the row; {@code false} to leave cells unmerged
     */
    public void addDiyRowContext(String context, boolean needMerge) { addDiyRowContext(context, needMerge, titleRowHeight); }

    /**
     * Appends a custom row with optional full-width merging and an explicit row height.
     *
     * @param context    the text to display
     * @param needMerge  {@code true} to merge all columns in the row; {@code false} to leave cells unmerged
     * @param h          row height in twips
     */
    public void addDiyRowContext(String context, boolean needMerge, int h) { addDiyRowContext(context, needMerge ? -1 : 0, needMerge ? -1 : 0, null, null, false, h); }

    // ==================== Freeze panes ====================

    /**
     * Freezes the top {@code rn} rows so they remain visible when scrolling vertically.
     * <p>Pass {@code 1} to freeze only the header row, {@code 2} to also freeze the title row, etc.
     *
     * @param rn the number of rows to freeze (1-based count from the top)
     */
    public void setFreezeRow(int rn) { setFreeze(0, rn, 0, rn); }

    /**
     * Freezes the leftmost {@code cn} columns so they remain visible when scrolling horizontally.
     *
     * @param cn the number of columns to freeze (1-based count from the left)
     */
    public void setFreezeColumn(int cn) { setFreeze(cn, 0, cn, 0); }

    /**
     * Freezes both the top {@code rn} rows and the leftmost {@code cn} columns simultaneously.
     *
     * @param rn the number of rows to freeze
     * @param cn the number of columns to freeze
     */
    public void setFreezeColumn(int rn, int cn) { setFreeze(cn, rn, cn, rn); }

    /**
     * Creates a freeze pane with full control over the split position and the top-left cell
     * of the scrollable region.
     *
     * @param sc the horizontal split position (number of columns frozen from the left)
     * @param sr the vertical split position (number of rows frozen from the top)
     * @param lc the left-most column index of the right scrollable region
     * @param er the top-most row index of the bottom scrollable region
     */
    public void setFreeze(int sc, int sr, int lc, int er) { sheet.createFreezePane(sc, sr, lc, er); }

    // ==================== Column management ====================

    /**
     * Removes the column identified by the given Java field name from the column mapping and header array,
     * then re-indexes all remaining columns to maintain a gap-free, sequential mapping.
     *
     * @param name the Java field name of the column to remove (must not be blank)
     * @throws NullPointerException if {@code name} is blank or null
     */
    public void removeExcelColumn(String name) {
        if (!Reflect.hasText(name)) throw new NullPointerException("column name is not null");
        for (Map.Entry<Integer, ExcelModel> entry : columnMappingInfo.entrySet()) {
            if (entry.getValue().getFieldName().equals(name)) { columnMappingInfo.remove(entry.getKey()); break; }
        }
        reorderHeader(); adjustExcelModelIndex();
    }

    /**
     * Changes the display order index of the column identified by the given field name.
     * <p>When {@code needReOrder} is {@code true} the header array and column mapping are
     * immediately re-sorted; set it to {@code false} when making several index changes in
     * a batch and calling {@link #reorderHeader()} manually afterwards.
     *
     * @param name        the Java field name of the column to reposition
     * @param index       the new display-order index (lower values appear further left)
     * @param needReOrder {@code true} to immediately re-sort and rebuild the header array
     */
    public void setHeaderIndex(String name, int index, boolean needReOrder) {
        for (Integer in : columnMappingInfo.keySet()) {
            if (columnMappingInfo.get(in).getFieldName().equals(name)) {
                columnMappingInfo.get(in).setIndex(index);
                if (needReOrder) reorderHeader();
                break;
            }
        }
    }

    /**
     * Sorts all columns by their {@link ExcelModel#getIndex()} value and rebuilds the
     * header array and column-mapping map to reflect the new order.
     * <p>Call this after any batch of {@link #setHeaderIndex} calls made with
     * {@code needReOrder = false}.
     */
    public void reorderHeader() {
        List<ExcelModel> sorted = columnMappingInfo.values().stream()
                .sorted((o1, o2) -> Integer.compare(o1.getIndex(), o2.getIndex()))
                .collect(Collectors.toList());
        header = new String[sorted.size()];
        columnMappingInfo.clear();
        for (int i = 0; i < sorted.size(); i++) {
            columnMappingInfo.put(i, sorted.get(i));
            header[i] = sorted.get(i).getColumnName();
        }
    }
    private void adjustExcelModelIndex() {
        columnNameModelMappingInfo.clear();
        AtomicInteger count = new AtomicInteger();
        columnMappingInfo.keySet().stream().sorted().forEach(key -> {
            ExcelModel m = columnMappingInfo.get(key);
            if (!Reflect.hasText(m.getColumnName()) && header != null && header.length > key)
                m.setColumnName(header[key]);
            if (m.getIndex() == 0) m.setIndex(key);
            m.setRealIndex(needOrderNum ? count.getAndIncrement() + 1 : count.getAndIncrement());
            columnNameModelMappingInfo.put(m.getFieldName(), m);
        });
    }

    // ==================== Static factory methods ====================

    /**
     * Creates an {@link ExcelModel} with full control over translation and picture flags.
     *
     * @param fieldName the Java field name to bind to this column
     * @param t         {@code true} if the cell value should be translated via a translation map
     * @param tm        the translation map ({@code value → display label}); used when {@code t} is {@code true}
     * @param p         {@code true} if this column contains picture URLs/paths
     * @return a new {@link ExcelModel} instance
     */
    public static ExcelModel generate(String fieldName, boolean t, Map<Object, Object> tm, boolean p) { return new ExcelModel(fieldName, t, tm, p); }

    /**
     * Creates a plain {@link ExcelModel} — no translation, no picture support.
     *
     * @param fieldName the Java field name to bind to this column
     * @return a new plain {@link ExcelModel} instance
     */
    public static ExcelModel generate(String fieldName) { return new ExcelModel(fieldName, false, null, false); }

    /**
     * Creates an {@link ExcelModel} that translates cell values using the supplied map.
     *
     * @param fieldName the Java field name to bind to this column
     * @param tm        the translation map ({@code value → display label})
     * @return a new translating {@link ExcelModel} instance
     */
    public static ExcelModel generate(String fieldName, Map<Object, Object> tm) { return new ExcelModel(fieldName, true, tm, false); }

    /**
     * Creates an {@link ExcelModel} that performs value-to-display-label translation using
     * separate property paths for the value and the display text.
     *
     * @param fieldName the Java field name to bind to this column
     * @param ivp       the "id/value" property path used to identify the raw value
     * @param idp       the "display" property path used to read the human-readable label
     * @return a new {@link ExcelModel} configured for property-path-based translation
     */
    public static ExcelModel generate(String fieldName, String ivp, String idp) { return new ExcelModel(fieldName, ivp, idp); }

    // ==================== Getters/Setters ====================

    void setExcelCreated(boolean v) { excelCreated = v; }
    boolean isChildComplex() { return isChildComplex; }
    PictureHandler getPictureHandler() { return pictureHandler; }
    DataValidator getDataValidator() { return dataValidator; }
    ExcelExporter getExporter() { return exporter; }
    Map<Integer, ExcelModel> getColumnMappingInfo() { return columnMappingInfo; }
    boolean isNeedOrderNum() { return needOrderNum; }
    int getRowNum() { return rowNum; }
    void setCurrentListNum(int n) { currentListNum = n; }

    /**
     * When {@code true}, a complex child creator inherits the header column count from
     * the immediately preceding sibling creator rather than computing it from its own header.
     *
     * @param v {@code true} to use the previous sibling's header length
     */
    public void setUsePrevHeaderLength(boolean v) { usePrevHeaderLength = v; }

    /**
     * When {@code true}, a complex child creator inherits the header column count from
     * the root (first) creator instead of computing its own.
     *
     * @param v {@code true} to use the root creator's header length
     */
    public void setUseRootHeaderLength(boolean v) { useRootHeaderLength = v; }

    /**
     * Sets the delimiter used to split multiple image URLs/paths stored in a single cell.
     * Defaults to {@code ","}.
     *
     * @param s the separator string (e.g. {@code "|"})
     */
    public void setImagesSeparator(String s) { imagesSeparator = s; }

    /**
     * Sets the height of the title row in twips (1/20th of a point). Defaults to {@code 1000}.
     *
     * @param h row height in twips
     */
    public void setTitleRowHeight(int h) { titleRowHeight = h; }

    /**
     * Sets the height of the header row in twips. Defaults to {@code 1000}.
     *
     * @param h row height in twips
     */
    public void setHeaderRowHeight(int h) { headerRowHeight = h; }

    /**
     * Sets the picture anchor type, controlling how embedded images move and resize with cells.
     * Use one of the {@code MOVE_AND_RESIZE}, {@code MOVE_DONT_RESIZE}, or
     * {@code DONT_MOVE_AND_RESIZE} constants.
     *
     * @param t the anchor type constant
     */
    public void setPictureType(int t) { pictureType = t; }

    /**
     * Sets the HTTP read timeout (in milliseconds) used when downloading remote image URLs.
     * Defaults to {@code 2000} ms.
     *
     * @param t timeout in milliseconds
     */
    public void setImageReadTimeOut(int t) { imageReadTimeOut = t; }

    /**
     * Sets the width of a specific column and records it in the internal width map.
     *
     * @param ci the zero-based column index
     * @param w  the column width in units of 1/256th of a character width (POI convention)
     */
    public void setColumnWidth(int ci, int w) { isSettingColumnWidth = true; columnWidthMap.put(ci, w); sheet.setColumnWidth(ci, w); }

    /**
     * When {@code true}, all columns not explicitly sized via {@link #setColumnWidth} are
     * auto-fitted to their content width after data population.
     *
     * <p>Width is <em>estimated</em> from the text length of a bounded sample of rows (up to
     * {@value #AUTOSIZE_SAMPLE_ROWS}) using a character-width heuristic — deliberately avoiding
     * POI's {@code Sheet.autoSizeColumn}, which renders every cell with font metrics (very slow
     * on large sheets) and throws on streaming sheets unless columns are tracked. The result is
     * approximate, not pixel-perfect.
     *
     * <p><strong>Big-data (streaming) mode:</strong> only rows still in the SXSSF window are
     * visible here, so the estimate is based on that window. For precise control on large
     * exports, set explicit widths via {@link #setColumnWidth} instead.
     *
     * @param autoSize {@code true} to enable auto-sizing
     */
    public void setAutoSizeColumns(boolean autoSize) { autoSizeColumns = autoSize; }

    /** Maximum number of rows sampled when estimating auto column widths. */
    private static final int AUTOSIZE_SAMPLE_ROWS = 1000;
    private static final int AUTOSIZE_MIN_WIDTH = 8 * 256;
    private static final int AUTOSIZE_MAX_WIDTH = 100 * 256;

    private void applyAutoSizeColumns() {
        if (!autoSizeColumns) return;
        int colCount = getHeaderLength();
        int firstRow = sheet.getFirstRowNum();
        int lastRow = Math.min(sheet.getLastRowNum(), firstRow + AUTOSIZE_SAMPLE_ROWS - 1);
        for (int c = 0; c < colCount; c++) {
            if (columnWidthMap.containsKey(c)) continue;
            int maxUnits = 0;
            for (int r = firstRow; r <= lastRow; r++) {
                Row sampleRow = sheet.getRow(r);
                if (sampleRow == null) continue;
                Cell sampleCell = sampleRow.getCell(c);
                if (sampleCell == null) continue;
                int units = displayWidth(cellText(sampleCell));
                if (units > maxUnits) maxUnits = units;
            }
            if (maxUnits == 0) continue;
            int width = Math.min(AUTOSIZE_MAX_WIDTH, Math.max(AUTOSIZE_MIN_WIDTH, (maxUnits + 2) * 256));
            sheet.setColumnWidth(c, width);
        }
    }

    /** Reads a cell's value as display text, tolerating any cell type. */
    private static String cellText(Cell cell) {
        try {
            return switch (cell.getCellType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                case FORMULA -> cell.getCellFormula();
                default -> "";
            };
        } catch (Exception e) {
            return "";
        }
    }

    /** Approximate display width in character units; wide (e.g. CJK) glyphs count as two. */
    private static int displayWidth(String s) {
        if (s == null) return 0;
        int w = 0;
        for (int i = 0; i < s.length(); i++) {
            w += s.charAt(i) > 0xFF ? 2 : 1;
        }
        return w;
    }

    /**
     * Sets a uniform height for all data rows. When {@code null} the default height
     * ({@code 1000} twips) is used.
     *
     * @param h row height in twips, or {@code null} to use the default
     */
    public void setRowHeight(Integer h) { rowHeight = h; }

    /**
     * Returns whether a sequential order-number column is prepended to every data row.
     *
     * @return {@code true} if the order-number column is enabled
     */
    public boolean getNeedOrderNum() { return needOrderNum; }

    /**
     * Enables or disables the automatic sequential order-number column prepended to each data row,
     * and re-indexes all {@link ExcelModel} instances to account for the extra column.
     *
     * @param v {@code true} to prepend an order-number column
     */
    public void setNeedOrderNum(boolean v) { needOrderNum = v; adjustExcelModelIndex(); }

    /**
     * Sets the number of columns that the order-number cell should span (merged horizontally).
     * Defaults to {@code 1} (no merge).
     *
     * @param span the column span count for the order-number cell
     */
    public void setOrderColumnSpan(int span) { orderColumnSpan = span; }

    /**
     * Sets the placeholder text written to cells whose resolved value is {@code null} or empty.
     *
     * @param v the default text (e.g. {@code "N/A"} or {@code "-"})
     */
    public void setNoneCellDefaultValue(String v) { noneCellDefaultValue = v; }

    /**
     * Returns the raw data object — either a {@link java.util.List} of row items
     * or a single {@link java.util.Map} row.
     *
     * @return the data source object
     */
    public Object getObject() { return object; }

    /**
     * Sets the data source object. Accepts a {@link java.util.List} of row objects
     * or a single {@link java.util.Map} for a one-row export.
     *
     * @param o the data source
     */
    public void setObject(Object o) { object = o; }

    /**
     * Returns the header label array in display order.
     *
     * @return the header string array, or {@code null} if not set
     */
    public String[] getHeader() { return header; }

    /**
     * Sets the header label array. Each element corresponds to one visible column.
     *
     * @param h the header string array
     */
    public void setHeader(String[] h) { header = h; }

    /**
     * Returns the title string displayed in the top merged row.
     *
     * @return the title, or {@code null} if no title was set
     */
    public String getTitle() { return title; }

    /**
     * Sets the title string. When non-blank a merged title row is written above the header.
     *
     * @param t the title text
     */
    public void setTitle(String t) { title = t; }

    /**
     * Returns the sheet name assigned to this creator's primary sheet.
     *
     * @return the sheet name
     */
    public String getSheetName() { return sheetName; }

    /**
     * Sets the sheet name for this creator's primary sheet.
     *
     * @param s the desired sheet name
     */
    public void setSheetName(String s) { sheetName = s; }

    /**
     * Replaces the entire column mapping and re-indexes all {@link ExcelModel} entries.
     * <p>The map key is the zero-based column position; the value is the {@link ExcelModel}
     * that describes the field binding, translation, and rendering for that column.
     *
     * @param m the new column mapping (key = column index, value = {@link ExcelModel})
     */
    public void setColumnMappingInfo(Map<Integer, ExcelModel> m) { columnMappingInfo = m; adjustExcelModelIndex(); }

    /**
     * Returns the vertical-merge configuration, mapping zero-based column indices to
     * the Java field names whose consecutive equal values trigger cell merging.
     *
     * @return the column-merge map
     */
    public Map<Integer, String> getColumnMergeInfo() { return columnMergeInfo; }

    /**
     * Sets the vertical-merge configuration.
     * <p>Each entry maps a zero-based column index to the field name of the bound data object;
     * consecutive rows with equal values in that field will be merged vertically.
     *
     * @param m the new column-merge map
     */
    public void setColumnMergeInfo(Map<Integer, String> m) { columnMergeInfo = m; }

    /**
     * Returns the list of child {@link ExcelCreator} instances that each generate an
     * additional sheet appended to this workbook.
     *
     * @return the mutable child creator list
     */
    public LinkedList<ExcelCreator> getChild() { return child; }

    /**
     * Replaces the list of child sheet creators.
     * <p>Each entry in the list generates one additional sheet in the same workbook
     * when {@link #createExcel()} is called.
     *
     * @param c the new list of child {@link ExcelCreator} instances
     */
    public void setChild(LinkedList<ExcelCreator> c) { child = c; }

    // ==================== Resource cleanup ====================

    /**
     * Releases resources held by this instance.
     * <p>Cleans up any temporary picture files managed by the {@link PictureHandler}.
     * When all active {@code ExcelCreator} instances have been closed, the shared
     * image-download thread pool is gracefully shut down (30-second timeout, then forced).
     * <p>It is safe to call this method more than once; subsequent calls after the first
     * are no-ops for the thread pool.
     */
    public void close() {
        if (pictureHandler != null) pictureHandler.cleanup();
        // GETTER_CACHE is a static shared cache; do not clear it on instance close to avoid affecting other concurrent instances
        if (INSTANCE_COUNT.decrementAndGet() <= 0) {
            synchronized (ExcelCreator.class) {
                if (executor != null && !executor.isShutdown()) {
                    executor.shutdown();
                    try {
                        if (!executor.awaitTermination(30, TimeUnit.SECONDS)) executor.shutdownNow();
                    } catch (InterruptedException ie) {
                        executor.shutdownNow();
                        Thread.currentThread().interrupt();
                    }
                    executor = null;
                }
            }
        }
        logger.debug("ExcelCreator resources closed");
    }
}
