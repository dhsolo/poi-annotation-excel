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
package io.github.dh.poi.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Auto-size estimates a formula column's width from its cached computed result, not from the
 * (potentially long) formula expression text.
 */
class CellTextFormulaTest {

    @Test
    void formulaCellMeasuredByCachedResultNotExpression() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            // A long expression whose freshly-cached numeric result is 0.0.
            cell.setCellFormula("1234567890+1234567890+1234567890+1234567890+1234567890");

            String text = ExcelCreator.cellText(cell);

            // Must reflect the cached value, never the long formula string.
            assertThat(text).doesNotContain("1234567890");
            assertThat(text.length()).isLessThan(8);
        }
    }

    @Test
    void plainCellTypesStillRead() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Row row = wb.createSheet().createRow(0);
            row.createCell(0).setCellValue("hello");
            row.createCell(1).setCellValue(true);

            assertThat(ExcelCreator.cellText(row.getCell(0))).isEqualTo("hello");
            assertThat(ExcelCreator.cellText(row.getCell(1))).isEqualTo("true");
        }
    }
}
