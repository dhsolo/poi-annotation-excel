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
package io.github.dhsolo.poi.excel;

import io.github.dhsolo.poi.excel.annotation.DefaultAnnotationProcessor;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import io.github.dhsolo.poi.excel.export.ExcelUploader;
import io.github.dhsolo.poi.excel.importor.ExcelImportor;
import io.github.dhsolo.poi.excel.importor.ExcelReadListener;
import io.github.dhsolo.poi.excel.annotation.ExcelColumn;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * Excel utility class — unified entry point covering all common export and import scenarios.
 *
 * <h3>Export — annotation model (simplest, recommended first)</h3>
 * <pre>
 * // 1. Define the model (only once)
 * {@literal @}ExcelInfo(sheetName = "Employees")
 * public class EmployeeExcel {
 *     {@literal @}ExcelTitle  private String title;
 *     {@literal @}ExcelData   private List{@literal <}Employee{@literal >} data;
 *     {@literal @}ExcelColumn(index = 0, columnName = "Name") private String name;
 *     {@literal @}ExcelColumn(index = 1, columnName = "Age")  private String age;
 * }
 *
 * // 2. Populate data and export (one line)
 * ExcelUtil.export(outputStream, "Employees", new EmployeeExcel().setTitle("Report").setData(list));
 * </pre>
 *
 * <h3>Export — Builder (no model class needed, suitable for dynamic columns)</h3>
 * <pre>
 * ExcelUtil.export(outputStream, "Report",
 *     ExcelCreatorBuilder.create("Sheet1")
 *         .title("Sales Report")
 *         .data(list)
 *         .columns("Name:name", "Amount:amount", "Date:date"));
 * </pre>
 *
 * <h3>Export — multiple sheets</h3>
 * <pre>
 * ExcelUtil.export(outputStream, "Summary",
 *     ExcelCreatorBuilder.create("Sales").data(salesList).columns("Name:name", "Amount:amount"),
 *     ExcelCreatorBuilder.create("Stock").data(stockList).columns("Goods:goods", "Qty:qty"));
 * </pre>
 *
 * <h3>Import (one line)</h3>
 * <pre>
 * List{@literal <}User{@literal >} users = ExcelUtil.importExcel(inputStream, User.class,
 *     ExcelModel.of("name"), ExcelModel.of("age"), ExcelModel.of("dept"));
 *
 * // When a POJO is not needed, get a Map directly
 * List{@literal <}Map{@literal <}String, Object{@literal >}{@literal >} rows =
 *     ExcelUtil.importExcelToMap(inputStream, ExcelModel.of("col1"), ExcelModel.of("col2"));
 * </pre>
 *
 * @author dh
 * @since 1.0
 */
public class ExcelUtil {

    // ==================== Export: annotation model ====================

    /**
     * Exports Excel to an output stream using an annotation model.
     * <p>The model class must be annotated with {@code @ExcelInfo}; data is carried via the {@code @ExcelData} field.
     *
     * @param out        target output stream
     * @param name       export file name (no extension needed)
     * @param excelModel annotation model instance carrying the data
     */
    public static void export(OutputStream out, String name, Object excelModel) throws IOException {
        ExcelCreator creator = new ExcelCreator(excelModel);
        try {
            creator.createExcel();
            creator.export(out, name);
        } finally {
            creator.close();
        }
    }

    /**
     * Exports Excel to an output stream using an annotation model, applying extra configuration before createExcel().
     * <p>Suitable for scenarios where custom rows or other modifications need to be added beyond the model.
     *
     * <pre>
     * ExcelUtil.export(out, model.getTitle(), model,
     *     creator -> { if (extraInfo != null) creator.addDiyRowContext(extraInfo, true); });
     * </pre>
     *
     * @param out          target output stream
     * @param name         export file name (no extension needed)
     * @param excelModel   annotation model instance carrying the data
     * @param configurator extra configuration applied before createExcel(); may be null
     */
    public static void export(OutputStream out, String name, Object excelModel,
                              Consumer<ExcelCreator> configurator) throws IOException {
        ExcelCreator creator = new ExcelCreator(excelModel);
        try {
            if (configurator != null) configurator.accept(creator);
            creator.createExcel();
            creator.export(out, name);
        } finally {
            creator.close();
        }
    }

    /**
     * Saves Excel to a local file using an annotation model.
     *
     * @param localPath  local save path (no extension needed)
     * @param excelModel annotation model instance carrying the data
     */
    public static void exportLocal(String localPath, Object excelModel) {
        ExcelCreator creator = new ExcelCreator(excelModel);
        try {
            creator.createExcel();
            creator.exportLocal(localPath);
        } finally {
            creator.close();
        }
    }

    /**
     * Generates Excel using an annotation model and returns it as an {@link InputStream},
     * suitable for uploading to OSS or similar scenarios.
     *
     * @param excelModel annotation model instance carrying the data
     * @return input stream of the Excel file
     */
    public static InputStream toInputStream(Object excelModel) {
        ExcelCreator creator = new ExcelCreator(excelModel);
        try {
            creator.createExcel();
            return creator.getInputStream();
        } finally {
            creator.close();
        }
    }

    /**
     * Generates Excel using an annotation model, applies extra configuration, then returns it as an
     * {@link InputStream}.
     *
     * @param excelModel   annotation model instance carrying the data
     * @param configurator extra configuration applied before createExcel(); may be {@code null}
     * @return input stream of the Excel file
     */
    public static InputStream toInputStream(Object excelModel, Consumer<ExcelCreator> configurator) {
        ExcelCreator creator = new ExcelCreator(excelModel);
        try {
            if (configurator != null) configurator.accept(creator);
            creator.createExcel();
            return creator.getInputStream();
        } finally {
            creator.close();
        }
    }

    /**
     * Generates Excel using an annotation model and returns the content as a {@code byte[]}.
     *
     * @param excelModel annotation model instance carrying the data
     * @return the serialised workbook as a byte array
     * @throws IOException if an I/O error occurs while serialising the workbook
     */
    public static byte[] toBytes(Object excelModel) throws IOException {
        try (InputStream in = toInputStream(excelModel)) {
            return in.readAllBytes();
        }
    }

    /**
     * Generates Excel using a fluent Builder chain and returns the content as a {@code byte[]}.
     *
     * @param builders one or more builders, each corresponding to one sheet
     * @return the serialised workbook as a byte array
     * @throws IOException if an I/O error occurs while serialising the workbook
     */
    public static byte[] toBytes(ExcelCreatorBuilder... builders) throws IOException {
        try (InputStream in = toInputStream(builders)) {
            return in.readAllBytes();
        }
    }

    // ==================== Export: Builder (skips the build→createExcel→export three steps) ====================

    /**
     * Exports directly to an output stream using a fluent Builder chain,
     * skipping the {@code build()-createExcel()-export()} three-step sequence.
     * <p>When multiple builders are provided, each corresponds to one sheet.
     *
     * <pre>
     * ExcelUtil.export(out, "report.xlsx",
     *     ExcelCreatorBuilder.create("Sheet1")
     *         .title("Sales Report")
     *         .data(list)
     *         .columns("Name:name", "Amount:amount"));
     * </pre>
     *
     * @param out      target output stream
     * @param name     export file name (no extension needed)
     * @param builders one or more builders, each corresponding to one sheet
     */
    public static void export(OutputStream out, String name, ExcelCreatorBuilder... builders) throws IOException {
        export(out, name, toCreators(builders));
    }

    /**
     * Saves to a local file using a fluent Builder chain,
     * skipping the {@code build()-createExcel()-exportLocal()} three-step sequence.
     * <p>When multiple builders are provided, each corresponds to one sheet.
     *
     * @param localPath local save path (no extension needed)
     * @param builders  one or more builders, each corresponding to one sheet
     */
    public static void exportLocal(String localPath, ExcelCreatorBuilder... builders) {
        exportLocal(localPath, toCreators(builders));
    }

    /**
     * Generates Excel using a fluent Builder chain and returns it as an {@link InputStream},
     * suitable for uploading to OSS or similar scenarios.
     *
     * @param builders one or more builders, each corresponding to one sheet
     * @return input stream of the Excel file
     */
    public static InputStream toInputStream(ExcelCreatorBuilder... builders) {
        if (builders == null || builders.length == 0) {
            throw new IllegalArgumentException("At least one Builder is required");
        }
        ExcelCreator[] creators = toCreators(builders);
        ExcelCreator first = creators[0];
        for (int i = 1; i < creators.length; i++) {
            first.getChild().add(creators[i]);
        }
        try {
            first.createExcel();
            return first.getInputStream();
        } finally {
            first.close();
        }
    }

    // ==================== Upload: annotation model ====================

    /**
     * Generates Excel using an annotation model and uploads it via {@link ExcelUploader}
     * to a custom storage backend (S3, MinIO, file service, etc.).
     *
     * <pre>
     * ExcelUtil.upload(
     *     (in, name) -> s3.putObject(bucket, name, in),
     *     "report.xlsx", model);
     * </pre>
     *
     * @param uploader   custom upload implementation
     * @param name       file name (extension recommended, e.g. {@code "report.xlsx"})
     * @param excelModel annotation model instance carrying the data
     * @param <R>        upload result type (e.g. URL string, file ID, etc.)
     * @return upload result returned by {@link ExcelUploader#upload}
     */
    public static <R> R upload(ExcelUploader<R> uploader, String name, Object excelModel) throws IOException {
        ExcelCreator creator = new ExcelCreator(excelModel);
        try {
            creator.createExcel();
            return creator.upload(uploader, name);
        } finally {
            creator.close();
        }
    }

    /**
     * Generates Excel using an annotation model, applies extra configuration, then uploads.
     *
     * @param uploader     custom upload implementation
     * @param name         file name
     * @param excelModel   annotation model instance carrying the data
     * @param configurator extra configuration applied before createExcel(); may be null
     * @param <R>          upload result type
     * @return upload result
     */
    public static <R> R upload(ExcelUploader<R> uploader, String name, Object excelModel,
                               Consumer<ExcelCreator> configurator) throws IOException {
        ExcelCreator creator = new ExcelCreator(excelModel);
        try {
            if (configurator != null) configurator.accept(creator);
            creator.createExcel();
            return creator.upload(uploader, name);
        } finally {
            creator.close();
        }
    }

    // ==================== Upload: Builder ====================

    /**
     * Generates Excel using a Builder and uploads it. Supports multiple sheets (each Builder = one sheet).
     *
     * <pre>
     * String url = ExcelUtil.upload(fileService::store, "report.xlsx",
     *     ExcelCreatorBuilder.create("Sheet1").data(list).columns("Name:name", "Amount:amount"));
     * </pre>
     *
     * @param uploader custom upload implementation
     * @param name     file name
     * @param builders one or more builders, each corresponding to one sheet
     * @param <R>      upload result type
     * @return upload result
     */
    public static <R> R upload(ExcelUploader<R> uploader, String name, ExcelCreatorBuilder... builders) throws IOException {
        return upload(uploader, name, toCreators(builders));
    }

    // ==================== Upload: ExcelCreator ====================

    /**
     * Generates Excel using pre-built {@link ExcelCreator} instances and uploads. Multiple instances = multiple sheets.
     *
     * @param uploader custom upload implementation
     * @param name     file name
     * @param creators one or more ExcelCreator instances, each corresponding to one sheet
     * @param <R>      upload result type
     * @return upload result
     */
    public static <R> R upload(ExcelUploader<R> uploader, String name, ExcelCreator... creators) throws IOException {
        int length = creators.length;
        if (length == 0) return null;
        ExcelCreator first = creators[0];
        for (int i = 1; i < length; i++) {
            first.getChild().add(creators[i]);
        }
        try {
            first.createExcel();
            return first.upload(uploader, name);
        } finally {
            first.close();
        }
    }

    // ==================== Export: ExcelCreator ====================

    /**
     * Exports to an output stream using pre-built {@link ExcelCreator} instances. Multiple instances = multiple sheets.
     *
     * @param out      target output stream
     * @param name     export file name (no extension needed)
     * @param creators one or more ExcelCreator instances, each corresponding to one sheet
     */
    public static void export(OutputStream out, String name, ExcelCreator... creators) throws IOException {
        int length = creators.length;
        if (length == 0) return;
        ExcelCreator first = creators[0];
        for (int i = 1; i < length; i++) {
            first.getChild().add(creators[i]);
        }
        try {
            first.createExcel();
            first.export(out, name);
        } finally {
            first.close();
        }
    }

    /**
     * Saves to a local file using pre-built {@link ExcelCreator} instances. Multiple instances = multiple sheets.
     *
     * @param localPath local save path (no extension needed)
     * @param creators  one or more ExcelCreator instances, each corresponding to one sheet
     */
    public static void exportLocal(String localPath, ExcelCreator... creators) {
        int length = creators.length;
        if (length == 0) return;
        ExcelCreator first = creators[0];
        for (int i = 1; i < length; i++) {
            first.getChild().add(creators[i]);
        }
        try {
            first.createExcel();
            first.exportLocal(localPath);
        } finally {
            first.close();
        }
    }

    // ==================== Import ====================

    /**
     * Imports data from the first sheet (index 0) of an Excel file, mapped to a list of the specified type.
     *
     * <p>Column mapping is resolved in one of two ways:
     * <ul>
     *   <li>If {@code columns} is non-empty, the supplied {@link ExcelModel} list is used as-is.</li>
     *   <li>If {@code columns} is empty and {@code clazz} is annotated with {@link ExcelInfo},
     *       columns are derived automatically from the class's {@code @ExcelColumn} fields,
     *       sorted by {@code @ExcelColumn.index()}. This is the recommended path when the same
     *       annotation model is used for both export and import.</li>
     * </ul>
     *
     * <pre>
     * // Manual mapping
     * List&lt;User&gt; users = ExcelUtil.importExcel(in, User.class,
     *     ExcelModel.of("name"), ExcelModel.of("age"), ExcelModel.of("dept"));
     *
     * // Annotation-driven (User is annotated with &#64;ExcelInfo / &#64;ExcelColumn)
     * List&lt;User&gt; users = ExcelUtil.importExcel(in, User.class);
     * </pre>
     *
     * @param in      Excel file input stream
     * @param clazz   target type — either annotated with {@link ExcelInfo} or supplied alongside explicit {@code columns}
     * @param columns column definitions in Excel column order; may be empty when {@code clazz} is annotated
     * @return list of parsed results
     * @throws IllegalStateException    when Excel parsing fails
     * @throws IllegalArgumentException when {@code columns} is empty and {@code clazz} has no usable {@link ExcelInfo} metadata
     */
    public static <T> List<T> importExcel(InputStream in, Class<T> clazz, ExcelModel... columns) {
        return importExcel(in, 0, clazz, columns);
    }

    /**
     * Imports data from the specified sheet, mapped to a list of the specified type.
     * <p>See {@link #importExcel(InputStream, Class, ExcelModel...)} for the column-resolution rules.
     *
     * @param in         Excel file input stream
     * @param sheetIndex sheet index (0-based)
     * @param clazz      target type — either annotated with {@link ExcelInfo} or supplied alongside explicit {@code columns}
     * @param columns    column definitions in Excel column order; may be empty when {@code clazz} is annotated
     * @return list of parsed results
     * @throws IllegalStateException    when Excel parsing fails
     * @throws IllegalArgumentException when {@code columns} is empty and {@code clazz} has no usable {@link ExcelInfo} metadata
     */
    public static <T> List<T> importExcel(InputStream in, int sheetIndex, Class<T> clazz, ExcelModel... columns) {
        ExcelImportor importor = new ExcelImportor(in);
        importor.addColumnName(resolveColumns(clazz, columns));
        boolean ok = importor.analysisExcel();
        if (!ok) {
            throw new IllegalStateException("Excel parsing failed: " + importor.getErrorMessage());
        }
        return importor.getObject(sheetIndex, clazz);
    }

    /**
     * Imports data from the first sheet (index 0) of an Excel file, returned as {@code List<Map>}; no POJO required.
     *
     * <pre>
     * List{@literal <}Map{@literal <}String, Object{@literal >}{@literal >} rows =
     *     ExcelUtil.importExcelToMap(inputStream,
     *         ExcelModel.of("col1"), ExcelModel.of("col2"), ExcelModel.of("col3"));
     * </pre>
     *
     * @param in      Excel file input stream
     * @param columns column definitions passed in Excel column order (use {@link ExcelModel#of} for quick creation)
     * @return each row represented as a Map, with field names as keys
     * @throws IllegalStateException when Excel parsing fails, with detailed error message
     */
    @SuppressWarnings("unchecked")
    public static List<Map<String, Object>> importExcelToMap(InputStream in, ExcelModel... columns) {
        return importExcelToMap(in, 0, columns);
    }

    /**
     * Imports data from the specified sheet of an Excel file, returned as {@code List<Map>}; no POJO required.
     *
     * @param in         Excel file input stream
     * @param sheetIndex sheet index (0-based)
     * @param columns    column definitions passed in Excel column order
     * @return each row represented as a Map, with field names as keys
     * @throws IllegalStateException when Excel parsing fails, with detailed error message
     */
    @SuppressWarnings("unchecked")
    public static List<Map<String, Object>> importExcelToMap(InputStream in, int sheetIndex, ExcelModel... columns) {
        ExcelImportor importor = new ExcelImportor(in);
        importor.addColumnName(toLinkedList(columns));
        boolean ok = importor.analysisExcel();
        if (!ok) {
            throw new IllegalStateException("Excel parsing failed: " + importor.getErrorMessage());
        }
        return (List<Map<String, Object>>) (List<?>) importor.getObject(sheetIndex, Map.class);
    }

    // ==================== Import: startRow variants ====================

    /**
     * Imports data from sheet {@code sheetIndex}, skipping the first {@code startRow} rows
     * (0-based). Use this when the Excel file has multi-row headers or leading metadata rows
     * that should not be treated as data.
     *
     * <pre>
     * // Skip 2 header rows (rows 0 and 1), read data starting from row 2
     * List&lt;User&gt; users = ExcelUtil.importExcel(in, 0, 2, User.class,
     *     ExcelModel.of("name"), ExcelModel.of("age"));
     * </pre>
     *
     * @param in         Excel file input stream
     * @param sheetIndex sheet index (0-based)
     * @param startRow   first row index to read as data (0-based; default is {@code 1})
     * @param clazz      target type
     * @param columns    column definitions in Excel column order
     * @return list of parsed results
     * @throws IllegalStateException when Excel parsing fails, with detailed error message
     */
    public static <T> List<T> importExcel(InputStream in, int sheetIndex, int startRow, Class<T> clazz, ExcelModel... columns) {
        ExcelImportor importor = new ExcelImportor(in);
        importor.addColumnName(resolveColumns(clazz, columns));
        importor.setStartRow(startRow);
        boolean ok = importor.analysisExcel();
        if (!ok) {
            throw new IllegalStateException("Excel parsing failed: " + importor.getErrorMessage());
        }
        return importor.getObject(sheetIndex, clazz);
    }

    /**
     * Imports data from sheet {@code sheetIndex} as {@code List<Map>}, skipping the first
     * {@code startRow} rows. Useful for files with multi-row headers or leading metadata.
     *
     * @param in         Excel file input stream
     * @param sheetIndex sheet index (0-based)
     * @param startRow   first row index to read as data (0-based; default is {@code 1})
     * @param columns    column definitions in Excel column order
     * @return each row represented as a Map, with field names as keys
     * @throws IllegalStateException when Excel parsing fails, with detailed error message
     */
    @SuppressWarnings("unchecked")
    public static List<Map<String, Object>> importExcelToMap(InputStream in, int sheetIndex, int startRow, ExcelModel... columns) {
        ExcelImportor importor = new ExcelImportor(in);
        importor.addColumnName(toLinkedList(columns));
        importor.setStartRow(startRow);
        boolean ok = importor.analysisExcel();
        if (!ok) {
            throw new IllegalStateException("Excel parsing failed: " + importor.getErrorMessage());
        }
        return (List<Map<String, Object>>) (List<?>) importor.getObject(sheetIndex, Map.class);
    }

    // ==================== Import: streaming (ExcelReadListener) ====================

    /**
     * Imports Excel data in streaming (row-by-row) mode from the first sheet, calling
     * {@link ExcelReadListener#onRow} for each parsed row and
     * {@link ExcelReadListener#onError} for each row-level validation or conversion error.
     *
     * <p>Unlike the batch variants, this method does not throw on validation errors;
     * instead, errors are delivered to the listener so that all rows are visited regardless.
     * Use {@link ExcelReadListener#onFinish} to detect overall completion.
     *
     * <pre>
     * ExcelUtil.importExcel(inputStream,
     *     new ExcelReadListener() {
     *         public void onRow(Map&lt;String, Object&gt; row, int idx) { process(row); }
     *         public void onError(String msg, int idx) { log.warn("row {} error: {}", idx, msg); }
     *     },
     *     ExcelModel.of("name"), ExcelModel.of("amount"));
     * </pre>
     *
     * @param in       Excel file input stream
     * @param listener row-level callback; receives each successfully parsed row and any errors
     * @param columns  column definitions in Excel column order
     */
    public static void importExcel(InputStream in, ExcelReadListener listener, ExcelModel... columns) {
        ExcelImportor importor = new ExcelImportor(in);
        importor.addColumnName(toLinkedList(columns));
        importor.setReadListener(listener);
        importor.analysisExcel(false);
    }

    /**
     * Streaming-import variant that derives column mappings from {@code clazz}'s
     * {@link ExcelInfo} / {@link ExcelColumn} annotations.
     * The listener still receives raw {@code Map} rows keyed by Java field name.
     *
     * <pre>
     * ExcelUtil.importExcel(inputStream, DeviceImportModel.class, new ExcelReadListener() {
     *     public void onRow(Map&lt;String, Object&gt; row, int idx) { process(row); }
     * });
     * </pre>
     *
     * @param in       Excel file input stream
     * @param clazz    annotated model class — must carry {@link ExcelInfo}
     * @param listener row-level callback
     * @throws IllegalArgumentException when {@code clazz} has no usable {@link ExcelInfo} metadata
     */
    public static <T> void importExcel(InputStream in, Class<T> clazz, ExcelReadListener listener) {
        ExcelImportor importor = new ExcelImportor(in);
        importor.addColumnName(resolveColumns(clazz, new ExcelModel[0]));
        importor.setReadListener(listener);
        importor.analysisExcel(false);
    }

    // ==================== Internal utilities ====================

    private static ExcelCreator[] toCreators(ExcelCreatorBuilder[] builders) {
        ExcelCreator[] creators = new ExcelCreator[builders.length];
        for (int i = 0; i < builders.length; i++) {
            creators[i] = builders[i].build();
        }
        return creators;
    }

    private static LinkedList<ExcelModel> toLinkedList(ExcelModel[] models) {
        LinkedList<ExcelModel> list = new LinkedList<>();
        for (ExcelModel m : models) {
            list.add(m);
        }
        return list;
    }

    /**
     * Returns the column list to drive import: explicit when {@code columns} is non-empty,
     * otherwise derived from {@code clazz}'s {@link ExcelInfo} annotation metadata.
     */
    private static <T> LinkedList<ExcelModel> resolveColumns(Class<T> clazz, ExcelModel[] columns) {
        if (columns != null && columns.length > 0) {
            return toLinkedList(columns);
        }
        if (clazz.getAnnotation(ExcelInfo.class) == null) {
            throw new IllegalArgumentException(
                    "No ExcelModel columns provided and class " + clazz.getName() +
                            " is not annotated with @ExcelInfo; cannot derive column mapping");
        }
        T template;
        try {
            template = clazz.getDeclaredConstructor().newInstance();
        } catch (ReflectiveOperationException e) {
            throw new IllegalArgumentException(
                    "Cannot derive columns from " + clazz.getName() +
                            ": class must have a public no-arg constructor", e);
        }
        DefaultAnnotationProcessor processor = new DefaultAnnotationProcessor(template);
        List<ExcelModel> models = processor.getExcelAnnotationProperty().getExcelModels();
        if (models == null || models.isEmpty()) {
            throw new IllegalArgumentException(
                    "Class " + clazz.getName() + " has no @ExcelColumn fields to derive column mapping");
        }
        return new LinkedList<>(models);
    }
}
