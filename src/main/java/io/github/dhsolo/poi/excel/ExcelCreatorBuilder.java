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

import io.github.dhsolo.poi.excel.picture.AnchorType;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.Map;

/**
 * Builder for constructing {@link ExcelCreator} instances with a fluent API.
 *
 * <p>Provides a convenient way to configure all Excel generation options
 * without dealing with multiple constructor overloads.
 *
 * <p>Typical usage:
 * <pre>
 * ExcelCreator creator = ExcelCreatorBuilder.create("Sheet1")
 *     .title("Report")
 *     .needOrderNum(true)
 *     .rowHeight(500)
 *     .build();
 * </pre>
 *
 * @author dhsolo
 * @since 1.0
 * @see ExcelCreator
 */
public class ExcelCreatorBuilder {

    private String sheetName;
    private String excelType = "xlsx";
    private boolean bigData = false;
    private String title;
    private String[] header;
    private Object data;
    private boolean needOrderNum = false;
    private int orderColumnSpan = 1;
    private Integer rowHeight;
    private Integer validationRowCount;
    private int titleRowHeight = 1000;
    private int headerRowHeight = 1000;
    private int pictureType = AnchorType.MOVE_AND_RESIZE.getValue();
    private int imageReadTimeOut = 2000;
    private String imagesSeparator = ",";
    private String noneCellDefaultValue;
    private Map<Integer, ExcelModel> columnMappingInfo = new LinkedHashMap<>();
    private Map<Integer, String> columnMergeInfo = new LinkedHashMap<>();
    private LinkedList<ExcelCreator> child = new LinkedList<>();
    private Map<Integer, Integer> columnWidthMap = new LinkedHashMap<>();
    private CellStyle titleCellStyle;
    private CellStyle cellStyle;
    private CellStyle headerCellStyle;
    private int freezeRows = 0;
    private int freezeCols = 0;
    private boolean autoSizeColumns = false;

    private ExcelCreatorBuilder(String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * Creates a new builder whose primary sheet will be named {@code sheetName}.
     *
     * @param sheetName the name of the first sheet; if blank, POI assigns a default name
     * @return a new {@code ExcelCreatorBuilder} instance
     */
    public static ExcelCreatorBuilder create(String sheetName) {
        return new ExcelCreatorBuilder(sheetName);
    }

    /**
     * Sets the Excel file format using a raw string value ({@code "xlsx"} or {@code "xls"}).
     * Prefer {@link #excelType(ExcelType)} for type safety.
     *
     * @param excelType the file format string
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder excelType(String excelType) {
        this.excelType = excelType;
        return this;
    }

    /**
     * Set the Excel file format using the {@link ExcelType} enum.
     *
     * @param excelType the Excel format type
     * @return this builder for method chaining
     * @see ExcelType
     */
    public ExcelCreatorBuilder excelType(ExcelType excelType) {
        this.excelType = excelType.getValue();
        return this;
    }

    /**
     * Enables streaming (SXSSFWorkbook) mode for large datasets that would otherwise exhaust heap memory.
     * Streaming mode does not support all POI features (e.g. reading back written cells).
     *
     * @param bigData {@code true} to activate big-data / streaming mode
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder bigData(boolean bigData) {
        this.bigData = bigData;
        return this;
    }

    /**
     * Sets the title text written in a merged row above the header.
     * When omitted no title row is emitted.
     *
     * @param title the title string
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder title(String title) {
        this.title = title;
        return this;
    }

    /**
     * Sets the header label array. Each element defines one column header cell.
     * When using {@link #columns(String...)} or {@link #columns(String[], String[])} this is set automatically.
     *
     * @param header array of column header labels in display order
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder header(String[] header) {
        this.header = header;
        return this;
    }

    /**
     * Sets the data source. Accepts a {@link java.util.List} of row objects or a single
     * {@link java.util.Map} for a one-row export.
     *
     * @param data the data to render into rows
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder data(Object data) {
        this.data = data;
        return this;
    }

    /**
     * Enables an automatic sequential order-number column prepended to every data row.
     *
     * @param needOrderNum {@code true} to prepend an order-number column
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder needOrderNum(boolean needOrderNum) {
        this.needOrderNum = needOrderNum;
        return this;
    }

    /**
     * Sets the number of columns that the order-number cell spans horizontally (merged).
     * Only relevant when {@link #needOrderNum(boolean)} is {@code true}. Defaults to {@code 1}.
     *
     * @param span the horizontal span of the order-number cell
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder orderColumnSpan(int span) {
        this.orderColumnSpan = span;
        return this;
    }

    /**
     * Sets a uniform height for all data rows in twips (1/20th of a point).
     * When not set, the default height ({@code 1000} twips) is used.
     *
     * @param rowHeight row height in twips
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder rowHeight(int rowHeight) {
        this.rowHeight = rowHeight;
        return this;
    }

    /**
     * Sets how many data rows (counted from the first data row) dropdown-list validations
     * and formula pre-fill cover. Defaults to {@code 1000}.
     *
     * @param rowCount the number of data rows to cover; must be &gt;= 1
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder validationRowCount(int rowCount) {
        this.validationRowCount = rowCount;
        return this;
    }

    /**
     * Sets the height of the title row in twips. Defaults to {@code 1000}.
     *
     * @param height title row height in twips
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder titleRowHeight(int height) {
        this.titleRowHeight = height;
        return this;
    }

    /**
     * Sets the height of the header row in twips. Defaults to {@code 1000}.
     *
     * @param height header row height in twips
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder headerRowHeight(int height) {
        this.headerRowHeight = height;
        return this;
    }

    /**
     * Sets the picture anchor type using a raw integer constant.
     * Prefer {@link #anchorType(AnchorType)} for type safety.
     * Valid values: {@link ExcelCreator#MOVE_AND_RESIZE}, {@link ExcelCreator#MOVE_DONT_RESIZE},
     * {@link ExcelCreator#DONT_MOVE_AND_RESIZE}.
     *
     * @param pictureType the anchor type constant
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder pictureType(int pictureType) {
        this.pictureType = pictureType;
        return this;
    }

    /**
     * Set the image anchor type using the {@link AnchorType} enum.
     *
     * @param anchorType the anchor behavior for images
     * @return this builder for method chaining
     * @see AnchorType
     */
    public ExcelCreatorBuilder anchorType(AnchorType anchorType) {
        this.pictureType = anchorType.getValue();
        return this;
    }

    /**
     * Sets the HTTP read timeout used when downloading remote image URLs.
     * Defaults to {@code 2000} ms.
     *
     * @param timeout timeout in milliseconds
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder imageReadTimeOut(int timeout) {
        this.imageReadTimeOut = timeout;
        return this;
    }

    /**
     * Sets the delimiter used to split multiple image URLs/paths stored in a single cell.
     * Defaults to {@code ","}.
     *
     * @param separator the separator string (e.g. {@code "|"})
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder imagesSeparator(String separator) {
        this.imagesSeparator = separator;
        return this;
    }

    /**
     * Sets the placeholder text written to cells whose resolved value is {@code null} or empty.
     *
     * @param value the default text (e.g. {@code "N/A"} or {@code "-"})
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder noneCellDefaultValue(String value) {
        this.noneCellDefaultValue = value;
        return this;
    }

    /**
     * Merges an existing column-mapping map into the builder's column mapping.
     * The map key is the zero-based column position; the value is the corresponding {@link ExcelModel}.
     *
     * @param mapping the column mapping to add (key = column index, value = {@link ExcelModel})
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder columnMapping(Map<Integer, ExcelModel> mapping) {
        this.columnMappingInfo.putAll(mapping);
        return this;
    }

    /**
     * Adds a single column definition at the specified position.
     *
     * @param index the zero-based column position
     * @param model the {@link ExcelModel} describing field binding and rendering for this column
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder addColumn(int index, ExcelModel model) {
        this.columnMappingInfo.put(index, model);
        return this;
    }

    /**
     * Declares column mappings in bulk using the format {@code "columnName:fieldName"}
     * (e.g. {@code "Name:name"}), and automatically sets the header array.
     * <p>Avoids the need for individual {@code addColumn} calls and a separate {@code header} call.
     *
     * <pre>
     * ExcelCreatorBuilder.create("Sheet1")
     *     .data(list)
     *     .columns("Name:name", "Age:age", "Department:dept");
     * </pre>
     *
     * @param columnDefs column definitions in "columnName:fieldName" format
     */
    public ExcelCreatorBuilder columns(String... columnDefs) {
        String[] headers = new String[columnDefs.length];
        for (int i = 0; i < columnDefs.length; i++) {
            String[] parts = columnDefs[i].split(":", 2);
            if (parts.length != 2) {
                throw new IllegalArgumentException(
                        "Invalid column definition format; expected 'columnName:fieldName', got: " + columnDefs[i]);
            }
            headers[i] = parts[0].trim();
            columnMappingInfo.put(i, new ExcelModel(parts[1].trim(), parts[0].trim(), i));
        }
        this.header = headers;
        return this;
    }

    /**
     * Declares column mappings in bulk where {@code headers} and {@code fields} correspond by position,
     * and automatically sets the header array.
     * <p>Suitable when the header array and field name array already exist separately.
     *
     * <pre>
     * ExcelCreatorBuilder.create("Sheet1")
     *     .data(list)
     *     .columns(new String[]{"Name","Age"}, new String[]{"name","age"});
     * </pre>
     *
     * @param headers display column name array
     * @param fields  corresponding Java field name array; must be the same length as headers
     */
    public ExcelCreatorBuilder columns(String[] headers, String[] fields) {
        if (headers.length != fields.length) {
            throw new IllegalArgumentException("headers and fields arrays must have the same length");
        }
        for (int i = 0; i < fields.length; i++) {
            columnMappingInfo.put(i, new ExcelModel(fields[i].trim(), headers[i].trim(), i));
        }
        this.header = headers;
        return this;
    }

    /**
     * Merges an existing vertical-merge configuration into the builder.
     * Each entry maps a zero-based column index to the Java field name whose consecutive
     * equal values will trigger automatic cell merging.
     *
     * @param mergeInfo the merge configuration to add
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder columnMerge(Map<Integer, String> mergeInfo) {
        this.columnMergeInfo.putAll(mergeInfo);
        return this;
    }

    /**
     * Adds a single vertical-merge rule: consecutive rows with equal values in {@code field}
     * will be merged in column {@code columnIndex}.
     *
     * @param columnIndex the zero-based column index to apply vertical merging to
     * @param field       the Java field name whose value determines merge boundaries
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder addMerge(int columnIndex, String field) {
        this.columnMergeInfo.put(columnIndex, field);
        return this;
    }

    /**
     * Replaces the entire list of child sheet creators.
     * Each child generates one additional sheet appended to the same workbook.
     *
     * @param child the list of child {@link ExcelCreator} instances
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder child(LinkedList<ExcelCreator> child) {
        this.child = child;
        return this;
    }

    /**
     * Appends a single child sheet creator to the builder's child list.
     *
     * @param childCreator the {@link ExcelCreator} to add as an additional sheet
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder addChild(ExcelCreator childCreator) {
        this.child.add(childCreator);
        return this;
    }

    /**
     * Sets the width of a specific column.
     *
     * @param columnIndex the zero-based column index
     * @param width       the column width in units of 1/256th of a character width (POI convention)
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder columnWidth(int columnIndex, int width) {
        this.columnWidthMap.put(columnIndex, width);
        return this;
    }

    /**
     * Freezes the top {@code rows} rows so they remain visible when scrolling vertically.
     * Applied to the sheet after {@link ExcelCreator#createExcel()} runs.
     * Pass {@code 0} to disable row freeze (default).
     *
     * @param rows number of rows to freeze from the top
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder freezeRow(int rows) {
        this.freezeRows = rows;
        return this;
    }

    /**
     * Freezes the leftmost {@code cols} columns so they remain visible when scrolling horizontally.
     * Applied to the sheet after {@link ExcelCreator#createExcel()} runs.
     * Pass {@code 0} to disable column freeze (default).
     *
     * @param cols number of columns to freeze from the left
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder freezeColumn(int cols) {
        this.freezeCols = cols;
        return this;
    }

    /**
     * Freezes both the top {@code rows} rows and the leftmost {@code cols} columns.
     *
     * @param rows number of rows to freeze (0 = none)
     * @param cols number of columns to freeze (0 = none)
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder freeze(int rows, int cols) {
        this.freezeRows = rows;
        this.freezeCols = cols;
        return this;
    }

    /**
     * Enables automatic column width fitting after all data rows have been written.
     * Columns whose width was set explicitly via {@link #columnWidth(int, int)} retain
     * their manual width; all other columns are auto-sized.
     * <p>Note: in streaming (big-data) mode only the rows still in the SXSSF window
     * are considered when computing widths.
     *
     * @param autoSize {@code true} to enable auto-sizing (default: {@code false})
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder autoSizeColumns(boolean autoSize) {
        this.autoSizeColumns = autoSize;
        return this;
    }

    /**
     * Overrides the default title cell style with a custom one.
     * Use {@link ExcelCreator#createCellStyle()} on the built creator to obtain a style object.
     *
     * @param style the custom {@link CellStyle} for title cells
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder titleCellStyle(CellStyle style) {
        this.titleCellStyle = style;
        return this;
    }

    /**
     * Overrides the default data cell style with a custom one.
     *
     * @param style the custom {@link CellStyle} for data cells
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder cellStyle(CellStyle style) {
        this.cellStyle = style;
        return this;
    }

    /**
     * Overrides the default header cell style with a custom one.
     *
     * @param style the custom {@link CellStyle} for header cells
     * @return this builder for method chaining
     */
    public ExcelCreatorBuilder headerCellStyle(CellStyle style) {
        this.headerCellStyle = style;
        return this;
    }

    /**
     * Builds and returns a fully configured {@link ExcelCreator} instance.
     * <p>All properties set on this builder are applied to the creator in the correct
     * initialization order. {@link ExcelCreator#createExcel()} must still be called
     * before exporting.
     *
     * @return a new, configured {@link ExcelCreator} ready for {@code createExcel()}
     */
    public ExcelCreator build() {
        ExcelCreator creator = new ExcelCreator(excelType, sheetName, bigData);
        if (title != null) creator.setTitle(title);
        if (header != null) creator.setHeader(header);
        if (data != null) creator.setObject(data);
        if (noneCellDefaultValue != null) creator.setNoneCellDefaultValue(noneCellDefaultValue);
        creator.setNeedOrderNum(needOrderNum);
        creator.setOrderColumnSpan(orderColumnSpan);
        creator.setTitleRowHeight(titleRowHeight);
        creator.setHeaderRowHeight(headerRowHeight);
        creator.setPictureType(pictureType);
        creator.setImageReadTimeOut(imageReadTimeOut);
        creator.setImagesSeparator(imagesSeparator);
        if (rowHeight != null) creator.setRowHeight(rowHeight);
        if (validationRowCount != null) creator.setValidationRowCount(validationRowCount);
        if (!columnMappingInfo.isEmpty()) creator.setColumnMappingInfo(columnMappingInfo);
        if (!columnMergeInfo.isEmpty()) creator.setColumnMergeInfo(columnMergeInfo);
        if (!child.isEmpty()) creator.setChild(child);
        if (titleCellStyle != null) creator.setTitleCellStyle(titleCellStyle);
        if (cellStyle != null) creator.setCellStyle(cellStyle);
        if (headerCellStyle != null) creator.setHeaderCellStyle(headerCellStyle);
        for (Map.Entry<Integer, Integer> entry : columnWidthMap.entrySet()) {
            creator.setColumnWidth(entry.getKey(), entry.getValue());
        }
        if (freezeRows > 0 || freezeCols > 0) {
            creator.setFreeze(freezeCols, freezeRows, freezeCols, freezeRows);
        }
        if (autoSizeColumns) {
            creator.setAutoSizeColumns(true);
        }
        return creator;
    }

    /**
     * Builds, creates, and returns the Excel content as an {@link InputStream}.
     * <p>Equivalent to {@code build().createExcel().getInputStream()} but in a single call.
     * Useful for directly streaming the result to an upload client.
     *
     * @return an {@link InputStream} over the serialised workbook bytes
     */
    public InputStream toInputStream() {
        ExcelCreator creator = build();
        try {
            creator.createExcel();
            return creator.getInputStream();
        } finally {
            creator.close();
        }
    }

    /**
     * Builds, creates, and returns the Excel content as a {@code byte[]}.
     * <p>Convenience alternative to {@link #toInputStream()} when an in-memory byte array
     * is preferred over a stream.
     *
     * @return the serialised workbook bytes
     * @throws IOException if an I/O error occurs while serialising the workbook
     */
    public byte[] toBytes() throws IOException {
        try (InputStream in = toInputStream()) {
            return in.readAllBytes();
        }
    }
}
