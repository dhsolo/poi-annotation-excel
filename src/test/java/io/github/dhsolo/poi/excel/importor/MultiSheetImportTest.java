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

import io.github.dhsolo.poi.excel.ExcelModel;
import io.github.dhsolo.poi.excel.ExcelUtil;
import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: importing from {@code sheetIndex > 0} used to register the column models at
 * index 0 regardless, so the Map variant returned one EMPTY map per row (silent data loss)
 * and the typed variant crashed with an {@code IndexOutOfBoundsException}.
 */
class MultiSheetImportTest {

    @Test
    void mapImportFromSecondSheetReturnsItsData() throws Exception {
        byte[] xlsx = twoSheetWorkbook();
        List<Map<String, Object>> rows = ExcelUtil.importExcelToMap(
                new ByteArrayInputStream(xlsx), 1, ExcelModel.of("name"));
        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).get("name")).isEqualTo("alpha");
    }

    @Test
    void typedImportFromSecondSheetReturnsItsData() throws Exception {
        byte[] xlsx = twoSheetWorkbook();
        List<Bean> rows = ExcelUtil.importExcel(
                new ByteArrayInputStream(xlsx), 1, 1, Bean.class);
        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getName()).isEqualTo("alpha");
    }

    @Test
    void firstSheetImportStillWorks() throws Exception {
        byte[] xlsx = twoSheetWorkbook();
        List<Map<String, Object>> rows = ExcelUtil.importExcelToMap(
                new ByteArrayInputStream(xlsx), 0, ExcelModel.of("code"));
        assertThat(rows).hasSize(2); // startRow defaults to 1: the two data rows of sheet 0
    }

    /** sheet0: code column with 3 rows (header+2 data); sheet1: 名称 header + "alpha". */
    private static byte[] twoSheetWorkbook() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet s0 = wb.createSheet("first");
            s0.createRow(0).createCell(0).setCellValue("编码");
            s0.createRow(1).createCell(0).setCellValue("c1");
            s0.createRow(2).createCell(0).setCellValue("c2");
            Sheet s1 = wb.createSheet("second");
            s1.createRow(0).createCell(0).setCellValue("名称");
            s1.createRow(1).createCell(0).setCellValue("alpha");
            wb.write(out);
            return out.toByteArray();
        }
    }

    @ExcelInfo(sheetName = "second")
    public static class Bean {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
    }
}
