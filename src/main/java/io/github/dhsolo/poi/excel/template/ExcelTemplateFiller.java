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
package io.github.dhsolo.poi.excel.template;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import io.github.dhsolo.common.Reflect;
import io.github.dhsolo.poi.excel.picture.ImageDownLoadTask;
import io.github.dhsolo.poi.excel.picture.PictureFormat;

import java.io.*;
import java.nio.file.Files;
import java.util.*;
import java.util.function.Function;
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
 * <p>Picture placeholders embed images instead of text. {@code ${@image:logo}} marks the cell
 * where a picture registered via {@link #fillPicture(String, byte[])} is anchored;
 * {@code ${list.@image:photo}} inside a list row takes the image from each row map's
 * {@code photo} value ({@code byte[]}, {@link File}, {@link InputStream}, or a {@link String}
 * URL / local path downloaded through the same guarded fetch as picture exports). The
 * placeholder text is cleared and the picture is anchored over the cell (moving and resizing
 * with it). The image format (PNG/JPEG/GIF/BMP) is detected from the bytes and preserved.
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
 * @author dhsolo
 * @since 1.0
 */
public class ExcelTemplateFiller {

    private static final Logger log = LoggerFactory.getLogger(ExcelTemplateFiller.class);
    private static final Pattern PLACEHOLDER = Pattern.compile("\\$\\{([^}]+)}");
    private static final Pattern IMAGE_PLACEHOLDER = Pattern.compile("\\$\\{@image:([^}]+)}");
    private static final int DEFAULT_IMAGE_READ_TIME_OUT = 2000;

    private final Workbook workbook;
    private final Map<String, Object> context = new LinkedHashMap<>();
    private final Map<String, List<Map<String, Object>>> listContext = new LinkedHashMap<>();
    private final Map<String, Object> pictureContext = new LinkedHashMap<>();
    /** URL/path → downloaded bytes; a {@code null} value caches a failure so it is logged once. */
    private final Map<String, byte[]> downloadCache = new HashMap<>();
    /** Bytes instance → media part index, so one image anchored many times is stored once. */
    private final Map<byte[], Integer> pictureIndexCache = new IdentityHashMap<>();
    private int imageReadTimeOut = DEFAULT_IMAGE_READ_TIME_OUT;
    private Pattern imagesSeparatorPattern = Pattern.compile(",");

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
     * Registers an image for a {@code ${@image:key}} placeholder. The placeholder cell's text
     * is cleared and the picture is anchored over the cell. The format (PNG/JPEG/GIF/BMP) is
     * detected from the bytes and preserved; unrecognized bytes are skipped with a warning.
     *
     * @param key   placeholder name, e.g. {@code "logo"} matches {@code ${@image:logo}}
     * @param image raw image bytes; {@code null} or empty leaves the placeholder to be cleared
     * @return this filler for method chaining
     */
    public ExcelTemplateFiller fillPicture(String key, byte[] image) {
        if (image != null && image.length > 0) pictureContext.put(key, image);
        return this;
    }

    /** Variant of {@link #fillPicture(String, byte[])} reading the image from a file. */
    public ExcelTemplateFiller fillPicture(String key, File image) throws IOException {
        return fillPicture(key, image == null ? null : Files.readAllBytes(image.toPath()));
    }

    /**
     * Variant of {@link #fillPicture(String, byte[])} reading the image from a stream.
     * The stream is fully read but not closed.
     */
    public ExcelTemplateFiller fillPicture(String key, InputStream image) throws IOException {
        return fillPicture(key, image == null ? null : image.readAllBytes());
    }

    /**
     * Variant of {@link #fillPicture(String, byte[])} that downloads the image from an
     * {@code http}/{@code https} URL (or reads a local file path) when the workbook is
     * written. The fetch goes through the same guards as picture exports — protocol
     * whitelist, {@link io.github.dhsolo.poi.excel.picture.ImageDownloadPolicy} (SSRF),
     * read timeout ({@link #imageReadTimeOut(int)}) and the size cap. A failed download
     * clears the placeholder with a warning instead of aborting the fill; identical
     * URLs are downloaded once and stored as a single media part.
     *
     * <p>The value may carry <b>multiple</b> URLs separated by {@link #imagesSeparator(String)}
     * (default {@code ","}, same as picture exports): images are anchored over the
     * placeholder cell and the cells to its right, one column per image.
     *
     * @param key      placeholder name, e.g. {@code "logo"} matches {@code ${@image:logo}}
     * @param imageUrl image URL(s) or local file path(s); blank values leave the placeholder to be cleared
     * @return this filler for method chaining
     */
    public ExcelTemplateFiller fillPicture(String key, String imageUrl) {
        if (imageUrl != null && !imageUrl.isBlank()) pictureContext.put(key, imageUrl);
        return this;
    }

    /**
     * Sets the connect/read timeout (milliseconds) for URL picture downloads.
     * Defaults to 2000 ms, matching the picture export default.
     */
    public ExcelTemplateFiller imageReadTimeOut(int millis) {
        this.imageReadTimeOut = millis;
        return this;
    }

    /**
     * Sets the separator splitting a {@code String} picture value into multiple URLs/paths.
     * Compiled as a regular expression (escape regex metacharacters, e.g. {@code "\\|"}),
     * exactly like the export-side {@code imagesSeparator}. Defaults to {@code ","}.
     */
    public ExcelTemplateFiller imagesSeparator(String separator) {
        this.imagesSeparatorPattern = Pattern.compile(separator != null && !separator.isEmpty() ? separator : ",");
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
        // Second pass: fill scalar placeholders (pictures first, so the leftover text
        // placeholders in the same cell are still substituted afterwards)
        for (Row row : sheet) {
            for (Cell cell : row) {
                fillImagePlaceholders(cell, IMAGE_PLACEHOLDER, pictureContext::get);
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
            // POI's shiftRows moves cells but NOT drawing anchors: pictures below the list
            // region would otherwise stay put and end up overlapping the inserted rows.
            shiftPictureAnchors(sheet, templateRowIndex, rows.size() - 1);
        }

        Pattern rowImagePattern = Pattern.compile("\\$\\{" + Pattern.quote(listKey) + "\\.@image:([^}]+)}");

        for (int i = 0; i < rows.size(); i++) {
            Row targetRow = (i == 0) ? templateRow : createRowFromSnapshot(sheet, templateRow, templateValues, templateRowIndex + i);
            Map<String, Object> rowData = rows.get(i);
            Map<String, Object> resolvedContext = new LinkedHashMap<>();
            for (Map.Entry<String, Object> entry : rowData.entrySet()) {
                resolvedContext.put(listKey + "." + entry.getKey(), entry.getValue());
            }
            for (Cell cell : targetRow) {
                fillImagePlaceholders(cell, rowImagePattern, rowData::get);
                fillCell(cell, resolvedContext);
            }
        }
    }

    /**
     * Shifts picture (and other shape) anchors located strictly below {@code insertPoint}
     * down by {@code shift} rows, mirroring what {@link Sheet#shiftRows} did to the cells.
     */
    private void shiftPictureAnchors(Sheet sheet, int insertPoint, int shift) {
        if (shift <= 0) return;
        Drawing<?> drawing = sheet.getDrawingPatriarch();
        if (drawing == null) return;
        for (Shape shape : drawing) {
            if (shape.getAnchor() instanceof ClientAnchor anchor && anchor.getRow1() > insertPoint) {
                anchor.setRow2(anchor.getRow2() + shift);
                anchor.setRow1(anchor.getRow1() + shift);
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

    /**
     * Replaces every image placeholder the given pattern finds in the cell with an empty string
     * and anchors the resolved picture over the cell. Group 1 of the pattern is the lookup key;
     * the resolver maps it to {@code byte[]}, {@link File}, {@link InputStream} or {@code null}
     * (unresolvable values just clear the placeholder).
     */
    private void fillImagePlaceholders(Cell cell, Pattern pattern, Function<String, Object> resolver) {
        if (cell.getCellType() != CellType.STRING) return;
        String raw = cell.getStringCellValue();
        if (!raw.contains("${")) return;

        Matcher m = pattern.matcher(raw);
        StringBuffer sb = new StringBuffer();
        boolean found = false;
        while (m.find()) {
            found = true;
            String name = m.group(1);
            int columnOffset = 0;
            for (byte[] data : toImageList(resolver.apply(name), name)) {
                if (insertPicture(cell, data, name, columnOffset)) {
                    columnOffset++;
                }
            }
            m.appendReplacement(sb, "");
        }
        if (found) {
            m.appendTail(sb);
            cell.setCellValue(sb.toString());
        }
    }

    /**
     * Resolves a picture placeholder value to the list of images it denotes. A {@code String}
     * is split on {@link #imagesSeparator(String)} and each part downloaded/read; other
     * supported types yield a single image. Unresolvable parts are skipped (logged in the
     * respective resolution step), so the list may be empty.
     */
    private List<byte[]> toImageList(Object value, String name) {
        if (value instanceof String urls) {
            List<byte[]> images = new ArrayList<>();
            for (String url : imagesSeparatorPattern.split(urls)) {
                byte[] data = downloadImage(url.trim(), name);
                if (data != null) images.add(data);
            }
            return images;
        }
        byte[] single = toImageBytes(value, name);
        return single == null ? List.of() : List.of(single);
    }

    /** Coerces a picture placeholder value to raw bytes; unusable values resolve to {@code null}. */
    private byte[] toImageBytes(Object value, String name) {
        try {
            if (value == null) return null;
            if (value instanceof byte[] bytes) return bytes.length > 0 ? bytes : null;
            if (value instanceof File file) return Files.readAllBytes(file.toPath());
            if (value instanceof InputStream in) return in.readAllBytes();
        } catch (IOException e) {
            log.warn("Image placeholder '{}' skipped: cannot read image source", name, e);
            return null;
        }
        log.warn("Image placeholder '{}' skipped: unsupported value type {}", name, value.getClass().getName());
        return null;
    }

    /**
     * Downloads (or reads, for a local path) image bytes through the guarded fetch shared
     * with picture exports. Results — including failures — are cached per URL, so the same
     * logo across a thousand list rows is fetched once and a dead URL times out once.
     */
    private byte[] downloadImage(String url, String name) {
        if (url.isBlank()) return null;
        if (downloadCache.containsKey(url)) return downloadCache.get(url);
        byte[] data = null;
        try (InputStream in = ImageDownLoadTask.openGuardedStream(url, imageReadTimeOut)) {
            data = in.readAllBytes();
        } catch (Exception e) {
            log.warn("Image placeholder '{}' skipped: download failed for {}", name, url, e);
        }
        downloadCache.put(url, data);
        return data;
    }

    /**
     * Anchors a picture over the placeholder cell shifted {@code columnOffset} columns to the
     * right (two-cell anchor spanning exactly one cell, moving and resizing with it) — offsets
     * beyond 0 host the extra images of a multi-URL value, mirroring the export behaviour.
     * The format is sniffed from the bytes so PNG transparency and the original encoding
     * survive; unrecognized bytes are skipped with a warning.
     *
     * @return {@code true} when a picture was anchored (the column slot is consumed)
     */
    private boolean insertPicture(Cell cell, byte[] data, String name, int columnOffset) {
        PictureFormat format = PictureFormat.sniff(data);
        if (format == null) {
            log.warn("Image placeholder '{}' skipped: unrecognized image format", name);
            return false;
        }
        Integer pictureIndex = pictureIndexCache.get(data);
        if (pictureIndex == null) {
            pictureIndex = workbook.addPicture(data, format.poiPictureType());
            pictureIndexCache.put(data, pictureIndex);
        }
        Sheet sheet = cell.getSheet();
        Drawing<?> drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            drawing = sheet.createDrawingPatriarch();
        }
        int column = cell.getColumnIndex() + columnOffset;
        ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(column);
        anchor.setRow1(cell.getRowIndex());
        anchor.setCol2(column + 1);
        anchor.setRow2(cell.getRowIndex() + 1);
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
        drawing.createPicture(anchor, pictureIndex);
        return true;
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
