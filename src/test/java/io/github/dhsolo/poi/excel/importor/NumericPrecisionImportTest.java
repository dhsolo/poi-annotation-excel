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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: numeric cells must keep their full precision on import.
 * The old implementation routed values through {@code NumberFormat.getInstance()},
 * which silently rounds to 3 fraction digits (0.123456789 became "0.123").
 */
class NumericPrecisionImportTest {

    @Test
    void numericCellKeepsAllFractionDigits() throws Exception {
        byte[] xlsx = workbookWithNumericCell(0.123456789);

        ExcelImportor importor = new ExcelImportor(new ByteArrayInputStream(xlsx));
        LinkedList<ExcelModel> models = new LinkedList<>();
        models.add(ExcelImportor.generateExcelModel("v"));
        importor.addColumnName(models);
        importor.setStartRow(1);

        assertThat(importor.analysisExcel()).isTrue();
        List<Map> rows = importor.getObject(0, Map.class);
        assertThat(rows.get(0).get("v").toString()).isEqualTo("0.123456789");
    }

    @Test
    void integerValuedNumericCellStaysPlain() throws Exception {
        byte[] xlsx = workbookWithNumericCell(12345.0);

        ExcelImportor importor = new ExcelImportor(new ByteArrayInputStream(xlsx));
        LinkedList<ExcelModel> models = new LinkedList<>();
        models.add(ExcelImportor.generateExcelModel("v"));
        importor.addColumnName(models);
        importor.setStartRow(1);

        assertThat(importor.analysisExcel()).isTrue();
        List<Map> rows = importor.getObject(0, Map.class);
        // No grouping separators and no trailing ".0"
        assertThat(rows.get(0).get("v").toString()).isEqualTo("12345");
    }

    /** NaN/Infinity cannot be represented as BigDecimal; formart must pass them through as text. */
    @Test
    void nanAndInfinityArePassedThroughAsText() throws Exception {
        ExcelImportor importor = new ExcelImportor(new ByteArrayInputStream(workbookWithNumericCell(1.0)));
        assertThat(importor.formart(Double.NaN)).isEqualTo("NaN");
        assertThat(importor.formart(Double.POSITIVE_INFINITY)).isEqualTo("Infinity");
        assertThat(importor.formart(Double.NEGATIVE_INFINITY)).isEqualTo("-Infinity");
    }

    private static byte[] workbookWithNumericCell(double value) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("Sheet1");
            sheet.createRow(0).createCell(0).setCellValue("v");
            sheet.createRow(1).createCell(0).setCellValue(value);
            wb.write(out);
            return out.toByteArray();
        }
    }
}
