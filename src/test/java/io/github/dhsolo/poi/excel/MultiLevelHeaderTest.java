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

import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import io.github.dhsolo.poi.excel.annotation.ExcelData;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/** Multi-level (3+ row) merged headers declared via {@code @ExcelColumn#groups()}. */
class MultiLevelHeaderTest {

    @Test
    void threeRowHeaderMergesGroupsAndLeaves() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new ThreeLevel().setData(List.of(new Row("A店", 1, 2, 3, 4))));

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet sheet = wb.getSheetAt(0);

            // Row 0: 门店 (col0, vertical-merged) | 2024年 (cols 1..4)
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("门店");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("2024年");
            // Row 1: 上半年 (cols 1..2) | 下半年 (cols 3..4)
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("上半年");
            assertThat(sheet.getRow(1).getCell(3).getStringCellValue()).isEqualTo("下半年");
            // Row 2: leaf names
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Q1");
            assertThat(sheet.getRow(2).getCell(4).getStringCellValue()).isEqualTo("Q4");
            // Data begins on row 3
            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("A店");
            assertThat(sheet.getRow(3).getCell(1).getNumericCellValue()).isEqualTo(1.0);

            assertThat(merges(sheet)).contains(
                    "0,2,0,0",   // 门店 vertical across the 3 header rows
                    "0,0,1,4",   // 2024年 across Q1..Q4
                    "1,1,1,2",   // 上半年 across Q1..Q2
                    "1,1,3,4");  // 下半年 across Q3..Q4
        }
    }

    @Test
    void fourRowHeaderSupportsArbitraryDepth() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new FourLevel().setData(List.of(new Row("B店", 5, 6, 7, 8))));

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet sheet = wb.getSheetAt(0);
            // 3 group rows + 1 leaf row = 4 header rows; data on row 4.
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("公司");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("2024年");
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("上半年");
            assertThat(sheet.getRow(3).getCell(1).getStringCellValue()).isEqualTo("Q1");
            assertThat(sheet.getRow(4).getCell(0).getStringCellValue()).isEqualTo("B店");

            assertThat(merges(sheet)).contains(
                    "0,3,0,0",   // 门店 leaf vertical across all 4 header rows
                    "0,0,1,4",   // 公司 across all four quarter columns
                    "1,1,1,4");  // 2024年 across all four quarter columns
        }
    }

    @Test
    void orderColumnSpansFullHeaderHeightAndShiftsData() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new OrderedThreeLevel().setData(List.of(new Row("A店", 1, 2, 3, 4))));

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet sheet = wb.getSheetAt(0);
            // 序号 at col0, 门店 at col1, quarters at cols 2..5 (everything shifted right by one).
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("序号");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("门店");
            assertThat(sheet.getRow(0).getCell(2).getStringCellValue()).isEqualTo("2024年");
            assertThat(sheet.getRow(2).getCell(2).getStringCellValue()).isEqualTo("Q1");
            // Data row 3: sequence number in col0, store in col1.
            assertThat(sheet.getRow(3).getCell(0).getNumericCellValue()).isEqualTo(1.0);
            assertThat(sheet.getRow(3).getCell(1).getStringCellValue()).isEqualTo("A店");

            assertThat(merges(sheet)).contains(
                    "0,2,0,0",   // 序号 spans all 3 header rows
                    "0,2,1,1",   // 门店 spans all 3 header rows
                    "0,0,2,5",   // 2024年 across the four quarter columns
                    "1,1,2,3",   // 上半年
                    "1,1,4,5");  // 下半年
        }
    }

    private static java.util.Set<String> merges(Sheet sheet) {
        java.util.Set<String> set = new java.util.HashSet<>();
        for (CellRangeAddress r : sheet.getMergedRegions()) {
            set.add(r.getFirstRow() + "," + r.getLastRow() + "," + r.getFirstColumn() + "," + r.getLastColumn());
        }
        return set;
    }

    @ExcelInfo(sheetName = "r", isBigData = false)
    public static class ThreeLevel {
        @ExcelData
        private List<Row> data;
        @ExcelColumn(columnName = "门店", index = 1)
        private String store;
        @ExcelColumn(columnName = "Q1", index = 2, groups = {"2024年", "上半年"})
        private Integer q1;
        @ExcelColumn(columnName = "Q2", index = 3, groups = {"2024年", "上半年"})
        private Integer q2;
        @ExcelColumn(columnName = "Q3", index = 4, groups = {"2024年", "下半年"})
        private Integer q3;
        @ExcelColumn(columnName = "Q4", index = 5, groups = {"2024年", "下半年"})
        private Integer q4;

        public ThreeLevel setData(List<Row> data) {
            this.data = data;
            return this;
        }

        public List<Row> getData() {
            return data;
        }
    }

    @ExcelInfo(sheetName = "r", isBigData = false)
    public static class FourLevel {
        @ExcelData
        private List<Row> data;
        @ExcelColumn(columnName = "门店", index = 1)
        private String store;
        @ExcelColumn(columnName = "Q1", index = 2, groups = {"公司", "2024年", "上半年"})
        private Integer q1;
        @ExcelColumn(columnName = "Q2", index = 3, groups = {"公司", "2024年", "上半年"})
        private Integer q2;
        @ExcelColumn(columnName = "Q3", index = 4, groups = {"公司", "2024年", "下半年"})
        private Integer q3;
        @ExcelColumn(columnName = "Q4", index = 5, groups = {"公司", "2024年", "下半年"})
        private Integer q4;

        public FourLevel setData(List<Row> data) {
            this.data = data;
            return this;
        }

        public List<Row> getData() {
            return data;
        }
    }

    @ExcelInfo(sheetName = "r", isBigData = false, needOrder = true)
    public static class OrderedThreeLevel {
        @ExcelData
        private List<Row> data;
        @ExcelColumn(columnName = "门店", index = 1)
        private String store;
        @ExcelColumn(columnName = "Q1", index = 2, groups = {"2024年", "上半年"})
        private Integer q1;
        @ExcelColumn(columnName = "Q2", index = 3, groups = {"2024年", "上半年"})
        private Integer q2;
        @ExcelColumn(columnName = "Q3", index = 4, groups = {"2024年", "下半年"})
        private Integer q3;
        @ExcelColumn(columnName = "Q4", index = 5, groups = {"2024年", "下半年"})
        private Integer q4;

        public OrderedThreeLevel setData(List<Row> data) {
            this.data = data;
            return this;
        }

        public List<Row> getData() {
            return data;
        }
    }

    public static class Row {
        private final String store;
        private final Integer q1, q2, q3, q4;

        public Row(String store, Integer q1, Integer q2, Integer q3, Integer q4) {
            this.store = store;
            this.q1 = q1;
            this.q2 = q2;
            this.q3 = q3;
            this.q4 = q4;
        }

        public String getStore() { return store; }
        public Integer getQ1() { return q1; }
        public Integer getQ2() { return q2; }
        public Integer getQ3() { return q3; }
        public Integer getQ4() { return q4; }
    }
}
