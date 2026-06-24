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

import io.github.dhsolo.poi.excel.ExcelUtil;
import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import io.github.dhsolo.poi.excel.annotation.ExcelDateFormat;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression coverage for the entity-import (DOM {@code getObject(clazz)}) fixes:
 * <ul>
 *   <li>all integral targets truncate decimals consistently (Long/Short/Byte, not just Integer);</li>
 *   <li>a cell that overflows its target type skips only that row, not the whole import;</li>
 *   <li>a column's own {@code @ExcelDateFormat} pattern is honoured when parsing back;</li>
 *   <li>fields without a setter are populated by direct field assignment.</li>
 * </ul>
 */
class EntityImportConversionTest {

    @Test
    void integralTargetsTruncateAndOverflowSkipsOnlyThatRow() throws Exception {
        byte[] xlsx = buildNumberSheet();

        List<NumModel> rows = ExcelUtil.importExcel(new ByteArrayInputStream(xlsx), NumModel.class);

        // Row 2 holds an Integer-overflowing value (9999999999) — it is dropped, the rest survive.
        assertThat(rows).hasSize(2);

        NumModel first = rows.get(0);
        assertThat(first.getI()).isEqualTo(3);        // 3.5 truncated toward zero
        assertThat(first.getL()).isEqualTo(7L);       // 7.9 truncated (was a NumberFormatException before)
        assertThat(first.getS()).isEqualTo((short) 2);// 2.0 -> 2
        assertThat(first.getB()).isEqualTo((byte) 1); // 1.0 -> 1

        NumModel second = rows.get(1);
        assertThat(second.getI()).isEqualTo(5);
        assertThat(second.getL()).isEqualTo(5_000_000_000L); // fits long, exceeds int
    }

    @Test
    void customDatePatternRoundTrips() throws Exception {
        byte[] xlsx = buildDateSheet(date(2024, Calendar.MAY, 6));

        List<DateModel> rows = ExcelUtil.importExcel(new ByteArrayInputStream(xlsx), DateModel.class);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getD()).isNotNull();
        assertThat(new SimpleDateFormat("yyyy-MM-dd").format(rows.get(0).getD())).isEqualTo("2024-05-06");
    }

    @Test
    void fieldsWithoutSetterAreAssignedDirectly() throws Exception {
        byte[] xlsx = buildTwoColumnTextSheet("Alice", "42");

        List<NoSetterModel> rows = ExcelUtil.importExcel(new ByteArrayInputStream(xlsx), NoSetterModel.class);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).name).isEqualTo("Alice"); // direct field assignment, no setter present
        assertThat(rows.get(0).age).isEqualTo(42);
    }

    // ---------- fixtures ----------

    private static byte[] buildNumberSheet() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("s");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("i");
            header.createCell(1).setCellValue("l");
            header.createCell(2).setCellValue("s");
            header.createCell(3).setCellValue("b");

            numericRow(sheet, 1, 3.5, 7.9, 2.0, 1.0);
            // Integer column overflows -> whole row must be skipped, not abort the import.
            numericRow(sheet, 2, 9_999_999_999.0, 1.0, 1.0, 1.0);
            numericRow(sheet, 3, 5.0, 5_000_000_000.0, 1.0, 1.0);

            wb.write(out);
            return out.toByteArray();
        }
    }

    private static void numericRow(Sheet sheet, int r, double i, double l, double s, double b) {
        Row row = sheet.createRow(r);
        row.createCell(0).setCellValue(i);
        row.createCell(1).setCellValue(l);
        row.createCell(2).setCellValue(s);
        row.createCell(3).setCellValue(b);
    }

    private static byte[] buildDateSheet(Date value) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("s");
            sheet.createRow(0).createCell(0).setCellValue("d");
            CreationHelper helper = wb.getCreationHelper();
            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(helper.createDataFormat().getFormat("yyyy-mm-dd"));
            Cell cell = sheet.createRow(1).createCell(0);
            cell.setCellValue(value);
            cell.setCellStyle(dateStyle);
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] buildTwoColumnTextSheet(String a, String b) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("s");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("name");
            header.createCell(1).setCellValue("age");
            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue(a);
            row.createCell(1).setCellValue(b);
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static Date date(int year, int month, int day) {
        Calendar c = Calendar.getInstance();
        c.clear();
        c.set(year, month, day);
        return c.getTime();
    }

    // ---------- models ----------

    @ExcelInfo(sheetName = "s")
    public static class NumModel {
        @ExcelColumn(columnName = "i", index = 1)
        private Integer i;
        @ExcelColumn(columnName = "l", index = 2)
        private Long l;
        @ExcelColumn(columnName = "s", index = 3)
        private Short s;
        @ExcelColumn(columnName = "b", index = 4)
        private Byte b;

        public Integer getI() { return i; }
        public void setI(Integer i) { this.i = i; }
        public Long getL() { return l; }
        public void setL(Long l) { this.l = l; }
        public Short getS() { return s; }
        public void setS(Short s) { this.s = s; }
        public Byte getB() { return b; }
        public void setB(Byte b) { this.b = b; }
    }

    @ExcelInfo(sheetName = "s")
    public static class DateModel {
        @ExcelColumn(columnName = "d", index = 1)
        @ExcelDateFormat(pattern = "dd-MM-yyyy")
        private Date d;

        public Date getD() { return d; }
        public void setD(Date d) { this.d = d; }
    }

    /** No setters on purpose: import must fall back to direct field assignment. */
    @ExcelInfo(sheetName = "s")
    public static class NoSetterModel {
        @ExcelColumn(columnName = "name", index = 1)
        public String name;
        @ExcelColumn(columnName = "age", index = 2)
        public Integer age;
    }
}
