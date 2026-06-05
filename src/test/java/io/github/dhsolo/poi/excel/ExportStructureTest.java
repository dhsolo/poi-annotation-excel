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
import io.github.dhsolo.poi.excel.annotation.ExcelListBox;
import io.github.dhsolo.poi.excel.annotation.ExcelRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/** Verifies that @ExcelListBox emits a dropdown data-validation and @ExcelRow emits a custom row. */
class ExportStructureTest {

    @Test
    void listBoxEmitsDataValidation() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new ListBoxModel().setData(List.of(new R("a"))));
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            XSSFSheet sheet = wb.getSheet("lb");
            assertThat(sheet.getDataValidations()).isNotEmpty();
        }
    }

    @Test
    void excelRowEmitsCustomRow() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new FootnoteModel().setData(List.of(new R("a"))));
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            assertThat(containsText(wb.getSheetAt(0), "脚注：仅供参考")).isTrue();
        }
    }

    private static boolean containsText(Sheet sheet, String text) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING
                        && text.equals(cell.getStringCellValue())) {
                    return true;
                }
            }
        }
        return false;
    }

    @ExcelInfo(sheetName = "lb", isBigData = false)
    public static class ListBoxModel {
        @ExcelData
        private List<R> data;

        @ExcelColumn(columnName = "状态", index = 1)
        @ExcelListBox(listTextBox = {"在线", "离线"})
        private String status;

        public ListBoxModel setData(List<R> data) { this.data = data; return this; }
        public List<R> getData() { return data; }
    }

    @ExcelInfo(sheetName = "fn", isBigData = false)
    public static class FootnoteModel {
        @ExcelData
        private List<R> data;

        @ExcelColumn(columnName = "状态", index = 1)
        private String status;

        @ExcelRow(order = 0)
        private String footnote = "脚注：仅供参考";

        public FootnoteModel setData(List<R> data) { this.data = data; return this; }
        public List<R> getData() { return data; }
        public String getFootnote() { return footnote; }
    }

    public static class R {
        private final String status;
        public R(String status) { this.status = status; }
        public String getStatus() { return status; }
    }
}
