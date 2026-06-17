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
import io.github.dhsolo.poi.excel.model.ComplexExcelModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A complex (multi-section, single-sheet) model whose child section uses a multi-level
 * {@code @ExcelColumn#groups()} header: the section is stacked below the parent section, and its
 * three-row merged header must render at the child's row offset (not at row 0).
 */
class ComplexMultiLevelHeaderTest {

    @Test
    void childSectionMultiLevelHeaderRendersAtItsOffset() throws Exception {
        Parent parent = new Parent();
        parent.prows = List.of(new PRow("总览"));
        MultiLevelChild child = new MultiLevelChild();
        child.rows = List.of(new DataRow("A店", 1, 2, 3, 4));
        parent.child = child;

        byte[] bytes = ExcelUtil.toBytes(parent);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet sheet = wb.getSheetAt(0);

            int parentHeaderRow = find(sheet, "名称");
            int topGroupRow = find(sheet, "2024年");          // child section's top group label

            // The child section (with its multi-level header) sits below the parent section.
            assertThat(parentHeaderRow).isGreaterThanOrEqualTo(0);
            assertThat(topGroupRow).isGreaterThan(parentHeaderRow);

            // The three header rows of the child line up from topGroupRow.
            int col = colOf(sheet, topGroupRow, "2024年");
            assertThat(sheet.getRow(topGroupRow + 1).getCell(col).getStringCellValue()).isEqualTo("上半年");
            assertThat(sheet.getRow(topGroupRow + 2).getCell(col).getStringCellValue()).isEqualTo("Q1");
            assertThat(sheet.getRow(topGroupRow + 2).getCell(col + 3).getStringCellValue()).isEqualTo("Q4");
            // Child data lands right below the three-row header.
            assertThat(sheet.getRow(topGroupRow + 3).getCell(col - 1).getStringCellValue()).isEqualTo("A店");

            // Merges are at the child's absolute offset, not row 0.
            assertThat(merges(sheet)).contains(
                    range(topGroupRow, topGroupRow + 2, col - 1, col - 1),   // 门店 vertical (3 rows)
                    range(topGroupRow, topGroupRow, col, col + 3),           // 2024年 across Q1..Q4
                    range(topGroupRow + 1, topGroupRow + 1, col, col + 1),   // 上半年
                    range(topGroupRow + 1, topGroupRow + 1, col + 2, col + 3)); // 下半年
        }
    }

    private static int find(Sheet sheet, String text) {
        for (Row r : sheet) {
            for (Cell c : r) {
                if (c.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING
                        && text.equals(c.getStringCellValue())) {
                    return r.getRowNum();
                }
            }
        }
        return -1;
    }

    private static int colOf(Sheet sheet, int row, String text) {
        for (Cell c : sheet.getRow(row)) {
            if (c.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING
                    && text.equals(c.getStringCellValue())) {
                return c.getColumnIndex();
            }
        }
        return -1;
    }

    private static String range(int r1, int r2, int c1, int c2) {
        return r1 + "," + r2 + "," + c1 + "," + c2;
    }

    private static java.util.Set<String> merges(Sheet sheet) {
        java.util.Set<String> set = new java.util.HashSet<>();
        for (CellRangeAddress r : sheet.getMergedRegions()) {
            set.add(range(r.getFirstRow(), r.getLastRow(), r.getFirstColumn(), r.getLastColumn()));
        }
        return set;
    }

    @ExcelInfo(sheetName = "complex", isBigData = false)
    public static class Parent implements ComplexExcelModel {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;
        @ExcelData
        public List<PRow> prows;

        MultiLevelChild child;

        @Override
        @SuppressWarnings("rawtypes")
        public List getComplexModels() {
            return List.of(child);
        }
    }

    @ExcelInfo(sheetName = "child", isBigData = false)
    public static class MultiLevelChild {
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
        @ExcelData
        public List<DataRow> rows;
    }

    public static class PRow {
        private final String name;

        public PRow(String name) {
            this.name = name;
        }

        public String getName() {
            return name;
        }
    }

    public static class DataRow {
        private final String store;
        private final Integer q1, q2, q3, q4;

        public DataRow(String store, Integer q1, Integer q2, Integer q3, Integer q4) {
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
