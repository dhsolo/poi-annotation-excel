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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class StreamingExcelReaderTest {

    @Test
    void streamsRowsAndMapsByHeader(@TempDir File tmp) throws Exception {
        File xlsx = new File(tmp, "data.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(); FileOutputStream fo = new FileOutputStream(xlsx)) {
            Sheet sheet = wb.createSheet("s1");
            Row h = sheet.createRow(0);
            h.createCell(0).setCellValue("名称");
            h.createCell(1).setCellValue("数量");
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("甲");
            r1.createCell(1).setCellValue(12);
            Row r2 = sheet.createRow(2);
            // leave column 0 empty to exercise gap-filling
            r2.createCell(1).setCellValue(7);
            wb.write(fo);
        }

        List<List<String>> rows = new ArrayList<>();
        try (FileInputStream in = new FileInputStream(xlsx)) {
            StreamingExcelReader.read(in, 0, (idx, cells) -> rows.add(cells));
        }
        assertThat(rows).hasSize(3);
        assertThat(rows.get(0)).containsExactly("名称", "数量");
        assertThat(rows.get(1)).containsExactly("甲", "12");
        assertThat(rows.get(2).get(0)).isEmpty();
        assertThat(rows.get(2).get(1)).isEqualTo("7");

        try (FileInputStream in = new FileInputStream(xlsx)) {
            List<Map<String, String>> maps = StreamingExcelReader.readAsMaps(in, 0);
            assertThat(maps).hasSize(2);
            assertThat(maps.get(0)).containsEntry("名称", "甲").containsEntry("数量", "12");
            assertThat(maps.get(1)).containsEntry("名称", "").containsEntry("数量", "7");
        }
    }

    @Test
    void mapsRowsToAnnotatedBeans(@TempDir File tmp) throws Exception {
        File xlsx = new File(tmp, "people.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(); FileOutputStream fo = new FileOutputStream(xlsx)) {
            Sheet sheet = wb.createSheet("p");
            Row h = sheet.createRow(0);
            h.createCell(0).setCellValue("名称");
            h.createCell(1).setCellValue("状态");
            h.createCell(2).setCellValue("数量");
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("甲");
            r1.createCell(1).setCellValue("在线");
            r1.createCell(2).setCellValue(12);
            Row r2 = sheet.createRow(2);
            r2.createCell(0).setCellValue("乙");
            r2.createCell(1).setCellValue("离线");
            r2.createCell(2).setCellValue(7);
            wb.write(fo);
        }

        try (FileInputStream in = new FileInputStream(xlsx)) {
            List<Person> people = StreamingExcelReader.readAsBeans(in, 0, Person.class);
            assertThat(people).hasSize(2);
            assertThat(people.get(0).name).isEqualTo("甲");
            assertThat(people.get(0).status).isEqualTo(1);   // 在线 -> 1 (reverse translate)
            assertThat(people.get(0).qty).isEqualTo(12);
            assertThat(people.get(1).name).isEqualTo("乙");
            assertThat(people.get(1).status).isEqualTo(0);   // 离线 -> 0
            assertThat(people.get(1).qty).isEqualTo(7);
        }
    }

    @Test
    void convertsDateColumns(@TempDir File tmp) throws Exception {
        File xlsx = new File(tmp, "dates.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(); FileOutputStream fo = new FileOutputStream(xlsx)) {
            Sheet sheet = wb.createSheet("d");
            CreationHelper helper = wb.getCreationHelper();
            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(helper.createDataFormat().getFormat("yyyy-mm-dd"));

            Row h = sheet.createRow(0);
            h.createCell(0).setCellValue("名称");
            h.createCell(1).setCellValue("日期");
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("甲");
            Cell dateCell = r1.createCell(1);          // real Excel date (serial number + date style)
            dateCell.setCellValue(java.sql.Date.valueOf("2024-06-01"));
            dateCell.setCellStyle(dateStyle);
            Row r2 = sheet.createRow(2);
            r2.createCell(0).setCellValue("乙");
            r2.createCell(1).setCellValue("2023-12-25");  // plain text date
            wb.write(fo);
        }

        try (FileInputStream in = new FileInputStream(xlsx)) {
            List<DateRow> rows = StreamingExcelReader.readAsBeans(in, 0, DateRow.class);
            assertThat(rows).hasSize(2);
            assertThat(rows.get(0).day).isEqualTo(LocalDate.of(2024, 6, 1));
            assertThat(rows.get(1).day).isEqualTo(LocalDate.of(2023, 12, 25));
        }
    }

    @Test
    void mapsByHeaderNameIgnoringOrderAndExtraColumns(@TempDir File tmp) throws Exception {
        File xlsx = new File(tmp, "shuffled.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(); FileOutputStream fo = new FileOutputStream(xlsx)) {
            Sheet sheet = wb.createSheet("p");
            // Columns in a different order than @ExcelColumn index, plus an unknown extra column.
            Row h = sheet.createRow(0);
            h.createCell(0).setCellValue("数量");
            h.createCell(1).setCellValue("备注");   // unknown column — must be ignored
            h.createCell(2).setCellValue("名称");
            h.createCell(3).setCellValue("状态");
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue(99);
            r1.createCell(1).setCellValue("x");
            r1.createCell(2).setCellValue("丙");
            r1.createCell(3).setCellValue("在线");
            wb.write(fo);
        }

        try (FileInputStream in = new FileInputStream(xlsx)) {
            List<Person> people = StreamingExcelReader.readAsBeans(in, 0, Person.class, true);
            assertThat(people).hasSize(1);
            assertThat(people.get(0).name).isEqualTo("丙");
            assertThat(people.get(0).status).isEqualTo(1);
            assertThat(people.get(0).qty).isEqualTo(99);
        }
    }

    @Test
    void skipsLeadingMetadataRowsViaHeaderRowIndex(@TempDir File tmp) throws Exception {
        File xlsx = new File(tmp, "withtitle.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(); FileOutputStream fo = new FileOutputStream(xlsx)) {
            Sheet sheet = wb.createSheet("p");
            sheet.createRow(0).createCell(0).setCellValue("人员报表");      // title row (ignored)
            Row h = sheet.createRow(1);                                     // header at index 1
            h.createCell(0).setCellValue("名称");
            h.createCell(1).setCellValue("状态");
            h.createCell(2).setCellValue("数量");
            Row r1 = sheet.createRow(2);
            r1.createCell(0).setCellValue("丁");
            r1.createCell(1).setCellValue("在线");
            r1.createCell(2).setCellValue(5);
            wb.write(fo);
        }

        try (FileInputStream in = new FileInputStream(xlsx)) {
            List<Person> people = StreamingExcelReader.readAsBeans(in, 0, Person.class, true, /*headerRowIndex=*/1);
            assertThat(people).hasSize(1);
            assertThat(people.get(0).name).isEqualTo("丁");
            assertThat(people.get(0).status).isEqualTo(1);
            assertThat(people.get(0).qty).isEqualTo(5);
        }
    }

    public static class Person {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelColumn(columnName = "状态", index = 2, translate = {"0:离线", "1:在线"})
        private Integer status;

        @ExcelColumn(columnName = "数量", index = 3)
        private int qty;
    }

    public static class DateRow {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelColumn(columnName = "日期", index = 2)
        private LocalDate day;
    }
}
