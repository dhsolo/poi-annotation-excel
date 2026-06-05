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
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/** Covers the dynamic Builder export path, multi-sheet export, and @ExcelColumn.sourceField. */
class ExcelExportMoreTest {

    @Test
    void builderExportsColumnsAndMapData() throws Exception {
        Map<String, Object> row = new LinkedHashMap<>();
        row.put("name", "甲");
        row.put("amount", 100);

        byte[] bytes = ExcelCreatorBuilder.create("S")
                .data(List.of(row))
                .columns("名称:name", "金额:amount")
                .bigData(false)
                .toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet s = wb.getSheetAt(0);
            assertThat(s.getRow(0).getCell(0).getStringCellValue()).isEqualTo("名称");
            assertThat(s.getRow(0).getCell(1).getStringCellValue()).isEqualTo("金额");
            assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("甲");
            assertThat(s.getRow(1).getCell(1).getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(s.getRow(1).getCell(1).getNumericCellValue()).isEqualTo(100.0);
        }
    }

    @Test
    void exportsMultipleSheets() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(
                ExcelCreatorBuilder.create("甲表").data(List.of(mapOf("name", "a"))).columns("名称:name").bigData(false),
                ExcelCreatorBuilder.create("乙表").data(List.of(mapOf("val", "b"))).columns("值:val").bigData(false));

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            // Both data sheets exist in one workbook (plus a hidden listConstantData helper sheet).
            assertThat(wb.getSheet("甲表")).isNotNull();
            assertThat(wb.getSheet("甲表").getRow(1).getCell(0).getStringCellValue()).isEqualTo("a");
            assertThat(wb.getSheet("乙表")).isNotNull();
            assertThat(wb.getSheet("乙表").getRow(1).getCell(0).getStringCellValue()).isEqualTo("b");
        }
    }

    private static Map<String, Object> mapOf(String k, Object v) {
        Map<String, Object> m = new LinkedHashMap<>();
        m.put(k, v);
        return m;
    }

    @Test
    void exportsValueViaSourceField() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new RedirectModel().setData(List.of(new Src("ok-value"))));
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet s = wb.getSheetAt(0);
            assertThat(s.getRow(0).getCell(0).getStringCellValue()).isEqualTo("结果");
            assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("ok-value");
        }
    }

    @ExcelInfo(sheetName = "redir", isBigData = false)
    public static class RedirectModel {
        @ExcelData
        private List<Src> data;

        // The display column "结果" reads its value from another field (correctResult).
        @ExcelColumn(columnName = "结果", index = 1, sourceField = "correctResult")
        private String result;

        public RedirectModel setData(List<Src> data) {
            this.data = data;
            return this;
        }

        public List<Src> getData() {
            return data;
        }
    }

    public static class Src {
        private final String correctResult;

        public Src(String correctResult) {
            this.correctResult = correctResult;
        }

        public String getCorrectResult() {
            return correctResult;
        }
    }
}
