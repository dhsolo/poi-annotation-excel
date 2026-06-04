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
package io.github.dh.poi.excel.template;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import io.github.dh.common.Reflect;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Fills an Excel template by replacing {@code ${fieldName}} placeholders with actual data.
 *
 * <p>Supports two fill modes:
 * <ul>
 *   <li><b>Single-object fill</b> — replaces placeholders with values from a {@code Map} or POJO.</li>
 *   <li><b>List fill</b> — expands a row tagged with {@code ${list.fieldName}} for each item in a list,
 *       inserting new rows and shifting the rest down.</li>
 * </ul>
 *
 * <p>Example template layout:
 * <pre>
 *   A1: Report Date  B1: ${reportDate}
 *   A3: No.          B3: Name          C3: Amount
 *   A4: ${list.no} B4: ${list.name} C4: ${list.amount}
 * </pre>
 *
 * <p>Example usage:
 * <pre>
 * ExcelTemplateFiller.of(templateStream)
 *     .fill("reportDate", "2024-01-01")
 *     .fillList("list", dataRows)
 *     .writeTo(outputStream);
 * </pre>
 *
 * @author dh
 * @since 1.0
 */
public class ExcelTemplateFiller {

    private static final Logger log = LoggerFactory.getLogger(ExcelTemplateFiller.class);
    private static final Pattern PLACEHOLDER = Pattern.compile("\\$\\{([^}]+)}");

    private final Workbook workbook;
    private final Map<String, Object> context = new LinkedHashMap<>();
    private final Map<String, List<Map<String, Object>>> listContext = new LinkedHashMap<>();

    private ExcelTemplateFiller(Workbook workbook) {
        this.workbook = workbook;
    }

    /** Creates a filler from an XLSX template input stream. */
    public static ExcelTemplateFiller of(InputStream templateStream) throws IOException {
        return new ExcelTemplateFiller(new XSSFWorkbook(templateStream));
    }

    /**
     * Creates a filler from an XLSX template supplied as a byte array.
     * Useful when the template is already in memory (e.g. loaded from a resource file or cache).
     *
     * @param template raw bytes of the XLSX template file
     * @return a new filler bound to the workbook parsed from the byte array
     * @throws IOException if the bytes cannot be parsed as a valid XLSX workbook
     */
    public static ExcelTemplateFiller of(byte[] template) throws IOException {
        return of(new ByteArrayInputStream(template));
    }

    /** Creates a filler from any existing {@link Workbook} instance. */
    public static ExcelTemplateFiller of(Workbook workbook) {
        return new ExcelTemplateFiller(workbook);
    }

    /** Adds a single key-value pair to the fill context. */
    public ExcelTemplateFiller fill(String key, Object value) {
        context.put(key, value);
        return this;
    }

    /** Adds all entries from the given map to the fill context. */
    public ExcelTemplateFiller fillAll(Map<String, Object> data) {
        if (data != null) context.putAll(data);
        return this;
    }

    /**
     * Adds a POJO to the fill context, using field names as keys.
     * Fields are accessed via reflection; {@code null} values are silently skipped.
     */
    public ExcelTemplateFiller fillBean(Object bean) {
        if (bean == null) return this;
        Class<?> cls = bean.getClass();
        Reflect.doWithFields(cls, field -> {
            field.setAccessible(true);
            try {
                Object val = field.get(bean);
                if (val != null) context.put(field.getName(), val);
            } catch (IllegalAccessException ignored) {}
        });
        return this;
    }

    /**
     * Registers a list of row data for list-fill mode.
     *
     * @param listKey  prefix used in placeholders, e.g. {@code "list"} matches {@code ${list.fieldName}}
     * @param rows     list of row maps; each map must use field names as keys
     */
    public ExcelTemplateFiller fillList(String listKey, List<Map<String, Object>> rows) {
        if (rows != null) listContext.put(listKey, rows);
        return this;
    }

    /**
     * Registers a list of POJOs for list-fill mode, converting each bean to a field-name→value map
     * via reflection. This avoids the boilerplate of manually building {@code List<Map>}.
     *
     * <pre>
     * filler.fillListBeans("list", orders);
     * // Equivalent to:
     * filler.fillList("list", orders.stream().map(o -> Map.of("id", o.getId(), "amount", o.getAmount())).toList());
     * </pre>
     *
     * <p>Fields with {@code null} values are included in the map as {@code null} entries so that
     * the placeholder is still replaced (with an empty string).
     *
     * @param listKey prefix used in placeholders, e.g. {@code "list"} matches {@code ${list.fieldName}}
     * @param beans   list of POJOs; may be {@code null} (treated as empty)
     * @param <T>     the bean type
     * @return this filler for method chaining
     */
    public <T> ExcelTemplateFiller fillListBeans(String listKey, List<T> beans) {
        if (beans == null || beans.isEmpty()) return this;
        List<Map<String, Object>> rows = new ArrayList<>(beans.size());
        for (T bean : beans) {
            Map<String, Object> row = new LinkedHashMap<>();
            Class<?> cls = bean.getClass();
            Reflect.doWithFields(cls, field -> {
                field.setAccessible(true);
                try {
                    row.put(field.getName(), field.get(bean));
                } catch (IllegalAccessException ignored) {}
            });
            rows.add(row);
        }
        listContext.put(listKey, rows);
        return this;
    }

    /**
     * Processes all sheets, performs substitution, and writes the result to {@code out}.
     *
     * @param out target output stream (not closed by this method)
     */
    public void writeTo(OutputStream out) throws IOException {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            processSheet(workbook.getSheetAt(i));
        }
        workbook.write(out);
    }

    /**
     * Convenience method that returns the filled workbook as a byte array.
     */
    public byte[] toBytes() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        writeTo(baos);
        return baos.toByteArray();
    }

    private void processSheet(Sheet sheet) {
        // First pass: handle list rows (they shift other rows, so do this before scalar fill)
        for (String listKey : listContext.keySet()) {
            expandListRows(sheet, listKey, listContext.get(listKey));
        }
        // Second pass: fill scalar placeholders
        for (Row row : sheet) {
            for (Cell cell : row) {
                fillCell(cell, context);
            }
        }
    }

    /**
     * Finds rows containing {@code ${listKey.}}, expands them for each data item,
     * and fills each expanded row with that item's values.
     */
    private void expandListRows(Sheet sheet, String listKey, List<Map<String, Object>> rows) {
        String prefix = "${" + listKey + ".";
        int templateRowIndex = -1;

        for (Row row : sheet) {
            for (Cell cell : row) {
                if (getCellStringValue(cell).contains(prefix)) {
                    templateRowIndex = row.getRowNum();
                    break;
                }
            }
            if (templateRowIndex >= 0) break;
        }

        if (templateRowIndex < 0) return;
        if (rows.isEmpty()) {
            // Clear placeholder cells so the template row is not left with raw ${...} text
            Row templateRow = sheet.getRow(templateRowIndex);
            if (templateRow != null) {
                for (Cell cell : templateRow) {
                    if (getCellStringValue(cell).contains(prefix)) {
                        cell.setCellValue("");
                    }
                }
            }
            return;
        }

        Row templateRow = sheet.getRow(templateRowIndex);

        // Snapshot template cell values before modifying the row in-place
        List<String> templateValues = new ArrayList<>();
        for (Cell cell : templateRow) {
            templateValues.add(getCellStringValue(cell));
        }

        // Insert (rows.size() - 1) new rows after the template row to make room
        if (rows.size() > 1) {
            sheet.shiftRows(templateRowIndex + 1, Math.max(sheet.getLastRowNum(), templateRowIndex + 1), rows.size() - 1);
        }

        for (int i = 0; i < rows.size(); i++) {
            Row targetRow = (i == 0) ? templateRow : createRowFromSnapshot(sheet, templateRow, templateValues, templateRowIndex + i);
            Map<String, Object> resolvedContext = new LinkedHashMap<>();
            for (Map.Entry<String, Object> entry : rows.get(i).entrySet()) {
                resolvedContext.put(listKey + "." + entry.getKey(), entry.getValue());
            }
            for (Cell cell : targetRow) {
                fillCell(cell, resolvedContext);
            }
        }
    }

    /**
     * Creates a new row using original template values (snapshotted before any fills),
     * preserving style from the source row.
     */
    private Row createRowFromSnapshot(Sheet sheet, Row source, List<String> templateValues, int destRowNum) {
        Row dest = sheet.createRow(destRowNum);
        dest.setHeight(source.getHeight());
        int col = 0;
        for (Cell srcCell : source) {
            Cell dstCell = dest.createCell(srcCell.getColumnIndex(), CellType.STRING);
            if (srcCell.getCellStyle() != null) {
                dstCell.setCellStyle(srcCell.getCellStyle());
            }
            if (col < templateValues.size()) {
                dstCell.setCellValue(templateValues.get(col));
            }
            col++;
        }
        return dest;
    }

    /** Replaces all {@code ${key}} placeholders in a cell's string value with context values. */
    private void fillCell(Cell cell, Map<String, Object> ctx) {
        if (cell.getCellType() != CellType.STRING) return;
        String raw = cell.getStringCellValue();
        if (!raw.contains("${")) return;

        Matcher m = PLACEHOLDER.matcher(raw);
        StringBuffer sb = new StringBuffer();
        boolean found = false;
        while (m.find()) {
            found = true;
            String key = m.group(1);
            Object val = ctx.get(key);
            m.appendReplacement(sb, val != null ? Matcher.quoteReplacement(val.toString()) : "");
        }
        m.appendTail(sb);
        if (found) {
            cell.setCellValue(sb.toString());
        }
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        return cell.getCellType() == CellType.STRING ? cell.getStringCellValue() : "";
    }
}
