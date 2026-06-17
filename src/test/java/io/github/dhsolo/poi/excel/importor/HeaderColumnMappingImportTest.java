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
import io.github.dhsolo.poi.excel.annotation.ExcelHeaderColumnMapping;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * {@code @ExcelHeaderColumnMapping}: a class-based import matches columns by header text (via the
 * method's {@code header -> field} map) rather than by position, so it is robust to column
 * reordering and to unknown extra columns in the source file.
 */
class HeaderColumnMappingImportTest {

    @Test
    void importsByHeaderNameRegardlessOfColumnOrder() throws Exception {
        // File order is REVERSED vs the model's index() order, with an extra unmapped column.
        byte[] xlsx = buildXlsx(new Object[][]{
                {"Unit Price", "Ignore Me", "Product Name"},   // header row
                {12.5, "junk", "Widget"},
                {3.0, "x", "Gadget"},
        });

        List<Product> rows = ExcelUtil.importExcel(new ByteArrayInputStream(xlsx), Product.class);

        assertThat(rows).hasSize(2);
        assertThat(rows.get(0).getProductName()).isEqualTo("Widget");
        assertThat(rows.get(0).getUnitPrice()).isEqualByComparingTo("12.5");
        assertThat(rows.get(1).getProductName()).isEqualTo("Gadget");
        assertThat(rows.get(1).getUnitPrice()).isEqualByComparingTo("3.0");
    }

    @Test
    void mappingThatMatchesNoHeaderYieldsEmptyResultInsteadOfThrowing() throws Exception {
        // Header texts are not in the mapping (e.g. wrong startRow / typo): every column is skipped,
        // so the import degrades to an empty result (and logs a WARN) rather than crashing.
        byte[] xlsx = buildXlsx(new Object[][]{
                {"AAA", "BBB"},
                {1.0, "x"},
        });

        List<Product> rows = ExcelUtil.importExcel(new ByteArrayInputStream(xlsx), Product.class);

        assertThat(rows).isEmpty();
    }

    private static byte[] buildXlsx(Object[][] data) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("s");
            for (int r = 0; r < data.length; r++) {
                Row row = sheet.createRow(r);
                for (int c = 0; c < data[r].length; c++) {
                    Object v = data[r][c];
                    if (v instanceof Number n) {
                        row.createCell(c).setCellValue(n.doubleValue());
                    } else {
                        row.createCell(c).setCellValue(String.valueOf(v));
                    }
                }
            }
            wb.write(out);
            return out.toByteArray();
        }
    }

    @ExcelInfo(sheetName = "s", isBigData = false)
    public static class Product {
        // index() order (productName, unitPrice) deliberately differs from the file's column order.
        @ExcelColumn(columnName = "名称", index = 1)
        private String productName;
        @ExcelColumn(columnName = "单价", index = 2)
        private BigDecimal unitPrice;

        @ExcelHeaderColumnMapping
        public Map<String, String> headerMapping() {
            Map<String, String> map = new LinkedHashMap<>();
            map.put("Product Name", "productName");
            map.put("Unit Price", "unitPrice");
            return map;
        }

        public String getProductName() { return productName; }
        public void setProductName(String productName) { this.productName = productName; }
        public BigDecimal getUnitPrice() { return unitPrice; }
        public void setUnitPrice(BigDecimal unitPrice) { this.unitPrice = unitPrice; }
    }
}
