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
import io.github.dhsolo.poi.excel.annotation.ExcelColumnParent;
import io.github.dhsolo.poi.excel.annotation.ExcelData;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/** Verifies @ExcelColumnParent renders a two-row grouped header and reads values via sourceField. */
class GroupedHeaderTest {

    @Test
    void rendersTwoRowGroupedHeader() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new GroupModel().setData(List.of(new Sale("甲", 10, 20))));

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet s = wb.getSheetAt(0);

            // Row 0 = parent header: "销售额" merged over the Q1/Q2 columns (1..2);
            // the non-grouped "名称" column shows its name on top and spans both header rows.
            assertThat(s.getRow(0).getCell(1).getStringCellValue()).isEqualTo("销售额");
            assertThat(s.getMergedRegions()).contains(new CellRangeAddress(0, 0, 1, 2));
            assertThat(s.getRow(0).getCell(0).getStringCellValue()).isEqualTo("名称");
            assertThat(s.getMergedRegions()).contains(new CellRangeAddress(0, 1, 0, 0));

            // Row 1 = column header row (sub-headers under the parent group).
            assertThat(s.getRow(1).getCell(1).getStringCellValue()).isEqualTo("Q1");
            assertThat(s.getRow(1).getCell(2).getStringCellValue()).isEqualTo("Q2");

            // Row 2 = data, read via sourceField q1/q2.
            assertThat(s.getRow(2).getCell(0).getStringCellValue()).isEqualTo("甲");
            assertThat(s.getRow(2).getCell(1).getNumericCellValue()).isEqualTo(10.0);
            assertThat(s.getRow(2).getCell(2).getNumericCellValue()).isEqualTo(20.0);
        }
    }

    @ExcelInfo(sheetName = "g", isBigData = false)
    public static class GroupModel {
        @ExcelData
        private List<Sale> data;

        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelColumnParent(value = "销售额", columns = {
                @ExcelColumn(columnName = "Q1", index = 2, sourceField = "q1"),
                @ExcelColumn(columnName = "Q2", index = 3, sourceField = "q2")
        })
        private Object sales;   // anchor field; values come from sourceField on the row

        public GroupModel setData(List<Sale> data) { this.data = data; return this; }
        public List<Sale> getData() { return data; }
    }

    public static class Sale {
        private final String name;
        private final Integer q1;
        private final Integer q2;

        public Sale(String name, Integer q1, Integer q2) {
            this.name = name;
            this.q1 = q1;
            this.q2 = q2;
        }

        public String getName() { return name; }
        public Integer getQ1() { return q1; }
        public Integer getQ2() { return q2; }
    }
}
