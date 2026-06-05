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
package io.github.dhsolo.poi.excel.importor;

import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Forward-only, low-memory reader for {@code .xlsx} tabular data, built on POI's event
 * (SAX) API ({@link XSSFReader} + {@link XSSFSheetXMLHandler}).
 *
 * <p>Unlike the DOM-based {@link ExcelImportor} (which loads the whole workbook into memory
 * via {@code WorkbookFactory.create}), this reader streams rows one at a time and never
 * materialises the full sheet, so it can read very large files with bounded memory — the
 * shared-string table aside.
 *
 * <p><strong>Scope / limitations</strong> (by design, to keep it additive and safe):
 * <ul>
 *   <li>{@code .xlsx} only (not {@code .xls} or CSV);</li>
 *   <li>tabular cell values only — no images, comments, or merged-region expansion;</li>
 *   <li>formula cells yield their last cached value;</li>
 *   <li>forward-only — there is no random row access.</li>
 * </ul>
 *
 * <p>The existing {@link ExcelImportor} is unchanged; use this class when memory matters more
 * than the richer feature set.
 *
 * @author dh
 * @since 1.0
 */
public final class StreamingExcelReader {

    private StreamingExcelReader() {}

    /**
     * Callback invoked once per row, in sheet order.
     */
    @FunctionalInterface
    public interface RowConsumer {
        /**
         * @param rowIndex zero-based row index within the sheet
         * @param cells    cell values in column order; gaps are filled with empty strings,
         *                  trailing empty cells may be omitted
         */
        void accept(int rowIndex, List<String> cells);
    }

    /**
     * Streams every row of the given sheet to {@code consumer}.
     *
     * @param xlsx       the workbook input stream; closed by this method
     * @param sheetIndex zero-based index of the sheet to read
     * @param consumer   receives each row as it is parsed
     * @throws Exception if the workbook cannot be opened or parsed
     */
    public static void read(InputStream xlsx, int sheetIndex, RowConsumer consumer) throws Exception {
        try (OPCPackage pkg = OPCPackage.open(xlsx)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings sst = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();

            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            int idx = 0;
            while (sheets.hasNext()) {
                try (InputStream sheetStream = sheets.next()) {
                    if (idx++ != sheetIndex) continue;
                    XMLReader parser = XMLHelper.newXMLReader();
                    parser.setContentHandler(new XSSFSheetXMLHandler(
                            styles, sst, new RowCollector(consumer), new DataFormatter(), false));
                    parser.parse(new InputSource(sheetStream));
                    return;
                }
            }
            throw new IllegalArgumentException("Sheet index out of range: " + sheetIndex);
        }
    }

    /**
     * Streams rows of the given sheet to {@code consumer}, skipping leading rows whose
     * zero-based index is below {@code startRowInclusive} (e.g. to ignore title/metadata rows).
     *
     * @param xlsx             the workbook input stream; closed by this method
     * @param sheetIndex       zero-based index of the sheet to read
     * @param startRowInclusive zero-based index of the first row to emit
     * @param consumer         receives each emitted row
     * @throws Exception if the workbook cannot be opened or parsed
     */
    public static void read(InputStream xlsx, int sheetIndex, int startRowInclusive, RowConsumer consumer)
            throws Exception {
        read(xlsx, sheetIndex, (rowIndex, cells) -> {
            if (rowIndex >= startRowInclusive) {
                consumer.accept(rowIndex, cells);
            }
        });
    }

    /**
     * Reads the sheet into a list of maps, using the first row as the header. Each subsequent
     * row becomes a {@code header -> value} map (column order preserved).
     *
     * @param xlsx       the workbook input stream; closed by this method
     * @param sheetIndex zero-based index of the sheet to read
     * @return one map per data row
     * @throws Exception if the workbook cannot be opened or parsed
     */
    public static List<Map<String, String>> readAsMaps(InputStream xlsx, int sheetIndex) throws Exception {
        return readAsMaps(xlsx, sheetIndex, 0);
    }

    /**
     * Reads the sheet into a list of maps. Rows before {@code headerRowIndex} are ignored; the
     * first row at or after it is used as the header; the rest become {@code header -> value}
     * maps (column order preserved). Use a non-zero {@code headerRowIndex} to skip leading
     * title/metadata rows.
     *
     * @param xlsx           the workbook input stream; closed by this method
     * @param sheetIndex     zero-based index of the sheet to read
     * @param headerRowIndex zero-based index at/after which the header row is located
     * @return one map per data row
     * @throws Exception if the workbook cannot be opened or parsed
     */
    public static List<Map<String, String>> readAsMaps(InputStream xlsx, int sheetIndex, int headerRowIndex)
            throws Exception {
        List<String> header = new ArrayList<>();
        boolean[] headerSeen = {false};
        List<Map<String, String>> rows = new ArrayList<>();
        read(xlsx, sheetIndex, (rowIndex, cells) -> {
            if (rowIndex < headerRowIndex) return;
            if (!headerSeen[0]) {
                headerSeen[0] = true;
                header.addAll(cells);
                return;
            }
            Map<String, String> row = new LinkedHashMap<>();
            for (int i = 0; i < header.size(); i++) {
                row.put(header.get(i), i < cells.size() ? cells.get(i) : "");
            }
            rows.add(row);
        });
        return rows;
    }

    /**
     * Streams the sheet and maps each data row to an instance of {@code type}, following the
     * same convention as {@link ExcelImportor}: columns bind to {@link ExcelColumn} fields by
     * position, ordered by {@link ExcelColumn#index()}. The first row is treated as the header
     * and skipped.
     *
     * <p>Value handling:
     * <ul>
     *   <li>{@link ExcelColumn#translate()} is applied in reverse (display text → stored value);</li>
     *   <li>cell text is converted to the field type for {@code String}, the boxed/primitive
     *       numeric types, {@link BigDecimal}, and {@code boolean};</li>
     *   <li>empty cells and unsupported field types are left at their default (the field is not set).</li>
     * </ul>
     *
     * @param xlsx       the workbook input stream; closed by this method
     * @param sheetIndex zero-based index of the sheet to read
     * @param type       the target bean type; must have a no-arg constructor
     * @param <T>        the bean type
     * @return one bean per data row
     * @throws Exception if the workbook cannot be opened/parsed or a row cannot be mapped
     */
    public static <T> List<T> readAsBeans(InputStream xlsx, int sheetIndex, Class<T> type) throws Exception {
        return readAsBeans(xlsx, sheetIndex, type, false);
    }

    /**
     * Variant of {@link #readAsBeans(InputStream, int, Class)} that can bind columns by header
     * name instead of by position.
     *
     * <p>When {@code matchByHeaderName} is {@code true}, each header cell is matched against the
     * column's {@link ExcelColumn#columnName()} (falling back to the field name when blank), so
     * column order in the file does not matter and extra/unknown columns are ignored. When
     * {@code false}, columns bind strictly by position ordered by {@link ExcelColumn#index()}.
     *
     * @param xlsx              the workbook input stream; closed by this method
     * @param sheetIndex        zero-based index of the sheet to read
     * @param type              the target bean type; must have a no-arg constructor
     * @param matchByHeaderName {@code true} to bind by header name, {@code false} to bind by position
     * @param <T>               the bean type
     * @return one bean per data row
     * @throws Exception if the workbook cannot be opened/parsed or a row cannot be mapped
     */
    public static <T> List<T> readAsBeans(InputStream xlsx, int sheetIndex, Class<T> type,
                                          boolean matchByHeaderName) throws Exception {
        return readAsBeans(xlsx, sheetIndex, type, matchByHeaderName, 0);
    }

    /**
     * Variant of {@link #readAsBeans(InputStream, int, Class, boolean)} that also skips leading
     * rows: rows before {@code headerRowIndex} are ignored and the first row at or after it is
     * the header. Use this to skip title/metadata rows above the header.
     *
     * @param xlsx              the workbook input stream; closed by this method
     * @param sheetIndex        zero-based index of the sheet to read
     * @param type              the target bean type; must have a no-arg constructor
     * @param matchByHeaderName {@code true} to bind by header name, {@code false} to bind by position
     * @param headerRowIndex    zero-based index at/after which the header row is located
     * @param <T>               the bean type
     * @return one bean per data row
     * @throws Exception if the workbook cannot be opened/parsed or a row cannot be mapped
     */
    public static <T> List<T> readAsBeans(InputStream xlsx, int sheetIndex, Class<T> type,
                                          boolean matchByHeaderName, int headerRowIndex) throws Exception {
        List<ColumnBinding> bindings = buildBindings(type);
        Map<String, ColumnBinding> byName = matchByHeaderName ? indexByName(bindings) : null;
        // In header-name mode this is populated from the header row: position -> binding (or null).
        List<ColumnBinding> resolved = new ArrayList<>();
        List<T> result = new ArrayList<>();
        boolean[] headerSeen = {false};
        read(xlsx, sheetIndex, (rowIndex, cells) -> {
            if (rowIndex < headerRowIndex) return;
            if (!headerSeen[0]) {
                headerSeen[0] = true;
                if (matchByHeaderName) {
                    for (String title : cells) {
                        resolved.add(byName.get(title == null ? "" : title.trim()));
                    }
                }
                return;
            }
            try {
                T bean = type.getDeclaredConstructor().newInstance();
                List<ColumnBinding> cols = matchByHeaderName ? resolved : bindings;
                int n = Math.min(cols.size(), cells.size());
                for (int i = 0; i < n; i++) {
                    ColumnBinding b = cols.get(i);
                    if (b == null) continue; // unmapped column
                    String raw = cells.get(i);
                    if (b.reverseTranslate != null && b.reverseTranslate.containsKey(raw)) {
                        raw = b.reverseTranslate.get(raw);
                    }
                    Object value = convert(raw, b.field.getType());
                    if (value != null) {
                        b.field.set(bean, value);
                    }
                }
                result.add(bean);
            } catch (Exception e) {
                throw new RuntimeException("Failed to map row " + rowIndex + " to " + type.getName(), e);
            }
        });
        return result;
    }

    /** Indexes bindings by their header display name for header-name matching. */
    private static Map<String, ColumnBinding> indexByName(List<ColumnBinding> bindings) {
        Map<String, ColumnBinding> byName = new HashMap<>();
        for (ColumnBinding b : bindings) {
            byName.putIfAbsent(b.displayName, b);
        }
        return byName;
    }

    /** Collects {@code @ExcelColumn} fields (including inherited) ordered by their index. */
    private static List<ColumnBinding> buildBindings(Class<?> type) {
        List<Field> fields = new ArrayList<>();
        for (Class<?> c = type; c != null && c != Object.class; c = c.getSuperclass()) {
            for (Field f : c.getDeclaredFields()) {
                if (f.isAnnotationPresent(ExcelColumn.class)) {
                    fields.add(f);
                }
            }
        }
        fields.sort(Comparator.comparingInt(f -> f.getAnnotation(ExcelColumn.class).index()));

        List<ColumnBinding> bindings = new ArrayList<>(fields.size());
        for (Field f : fields) {
            f.setAccessible(true);
            ExcelColumn col = f.getAnnotation(ExcelColumn.class);
            String displayName = col.columnName().isEmpty() ? f.getName() : col.columnName().trim();
            Map<String, String> reverse = null;
            String[] translate = col.translate();
            if (translate.length > 0) {
                reverse = new HashMap<>();
                for (String t : translate) {
                    int sep = t.indexOf(':');
                    if (sep >= 0) {
                        reverse.put(t.substring(sep + 1), t.substring(0, sep)); // display -> stored
                    }
                }
            }
            bindings.add(new ColumnBinding(f, displayName, reverse));
        }
        return bindings;
    }

    /** Converts cell text to the target field type; returns {@code null} when not convertible. */
    private static Object convert(String raw, Class<?> target) {
        if (raw == null || raw.isEmpty()) return null;
        if (target == String.class) return raw;
        String s = raw.trim();
        boolean integral = s.indexOf('.') < 0 && s.indexOf('e') < 0 && s.indexOf('E') < 0;
        try {
            // Parse integral text directly (Long.parseLong etc.) so large values keep full
            // precision; only route through double when the text has a fractional/exponent part.
            if (target == Integer.class || target == int.class) return integral ? Integer.parseInt(s) : (int) Double.parseDouble(s);
            if (target == Long.class || target == long.class) return integral ? Long.parseLong(s) : (long) Double.parseDouble(s);
            if (target == Double.class || target == double.class) return Double.parseDouble(s);
            if (target == Float.class || target == float.class) return Float.parseFloat(s);
            if (target == Short.class || target == short.class) return integral ? Short.parseShort(s) : (short) Double.parseDouble(s);
            if (target == Byte.class || target == byte.class) return integral ? Byte.parseByte(s) : (byte) Double.parseDouble(s);
            if (target == BigDecimal.class) return new BigDecimal(s);
            if (target == Boolean.class || target == boolean.class) return parseBoolean(s);
        } catch (NumberFormatException e) {
            return null;
        }
        if (target == LocalDate.class || target == LocalDateTime.class
                || target == LocalTime.class || target == Date.class) {
            return convertTemporal(s, target);
        }
        return null; // unsupported type — left at default
    }

    private static final DateTimeFormatter[] DATE_TIME_PATTERNS = {
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss"),
            DateTimeFormatter.ofPattern("yyyy/M/d H:mm:ss"),
            DateTimeFormatter.ofPattern("yyyy/M/d H:mm"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm"),
    };
    private static final DateTimeFormatter[] DATE_PATTERNS = {
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            DateTimeFormatter.ofPattern("yyyy/M/d"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd"),
            DateTimeFormatter.ofPattern("yyyyMMdd"),
    };
    private static final DateTimeFormatter[] TIME_PATTERNS = {
            DateTimeFormatter.ofPattern("HH:mm:ss"),
            DateTimeFormatter.ofPattern("H:mm"),
    };

    /**
     * Converts cell text to a date/time field. Pure-numeric text is treated as an Excel serial
     * date; otherwise a set of common date, date-time, and time patterns is tried. Returns
     * {@code null} when nothing parses.
     */
    private static Object convertTemporal(String s, Class<?> target) {
        LocalDateTime ldt = null;
        try {
            ldt = DateUtil.getLocalDateTime(Double.parseDouble(s)); // Excel serial date
        } catch (NumberFormatException ignored) {
            for (DateTimeFormatter f : DATE_TIME_PATTERNS) {
                try { ldt = LocalDateTime.parse(s, f); break; } catch (Exception ignore) {}
            }
            if (ldt == null) {
                for (DateTimeFormatter f : DATE_PATTERNS) {
                    try { ldt = LocalDate.parse(s, f).atStartOfDay(); break; } catch (Exception ignore) {}
                }
            }
            if (ldt == null && target == LocalTime.class) {
                for (DateTimeFormatter f : TIME_PATTERNS) {
                    try { return LocalTime.parse(s, f); } catch (Exception ignore) {}
                }
            }
        } catch (Exception e) {
            return null;
        }
        if (ldt == null) return null;
        if (target == LocalDateTime.class) return ldt;
        if (target == LocalDate.class) return ldt.toLocalDate();
        if (target == LocalTime.class) return ldt.toLocalTime();
        if (target == Date.class) return Date.from(ldt.atZone(ZoneId.systemDefault()).toInstant());
        return null;
    }

    private static boolean parseBoolean(String s) {
        return s.equalsIgnoreCase("true") || s.equals("1")
                || s.equalsIgnoreCase("yes") || s.equalsIgnoreCase("y") || s.equals("是");
    }

    /** A column-to-field binding with its header name and an optional reverse translation map. */
    private static final class ColumnBinding {
        final Field field;
        final String displayName;
        final Map<String, String> reverseTranslate;

        ColumnBinding(Field field, String displayName, Map<String, String> reverseTranslate) {
            this.field = field;
            this.displayName = displayName;
            this.reverseTranslate = reverseTranslate;
        }
    }

    /**
     * Bridges POI's SAX {@link SheetContentsHandler} to the {@link RowConsumer}, filling cell
     * gaps so that each row's list is positional by column.
     */
    private static final class RowCollector implements SheetContentsHandler {

        private final RowConsumer consumer;
        private List<String> current;
        private int nextCol;

        RowCollector(RowConsumer consumer) {
            this.consumer = consumer;
        }

        @Override
        public void startRow(int rowNum) {
            current = new ArrayList<>();
            nextCol = 0;
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            int col = cellReference == null ? nextCol : new CellReference(cellReference).getCol();
            while (nextCol < col) {
                current.add("");
                nextCol++;
            }
            current.add(formattedValue == null ? "" : formattedValue);
            nextCol++;
        }

        @Override
        public void endRow(int rowNum) {
            consumer.accept(rowNum, current);
        }
    }
}
