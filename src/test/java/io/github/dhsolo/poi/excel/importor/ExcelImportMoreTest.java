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
import io.github.dhsolo.poi.excel.annotation.ExcelData;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/** Round-trips annotation export → annotation import, and covers importExcelToMap. */
class ExcelImportMoreTest {

    @Test
    void annotationExportThenImportRoundTrip() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new ExportModel().setData(List.of(
                new Dev("路由器", "网络部"),
                new Dev("交换机", "网络部"))));

        // Round-trip via the streaming reader (binds by @ExcelColumn position, first row = header).
        List<ImportRow> rows = StreamingExcelReader.readAsBeans(new ByteArrayInputStream(bytes), 0, ImportRow.class);

        assertThat(rows).hasSize(2);
        assertThat(rows.get(0).getName()).isEqualTo("路由器");
        assertThat(rows.get(0).getDept()).isEqualTo("网络部");
        assertThat(rows.get(1).getName()).isEqualTo("交换机");
    }

    @Test
    void importToMapWithManualColumns() throws Exception {
        byte[] xlsx = buildXlsx(new String[][] {
                {"姓名", "部门"},
                {"Alice", "研发"},
                {"Bob", "测试"},
        });

        List<Map<String, Object>> rows = ExcelUtil.importExcelToMap(
                new ByteArrayInputStream(xlsx), 0, 1, ExcelModel.of("name"), ExcelModel.of("dept"));

        assertThat(rows).hasSize(2);
        assertThat(rows.get(0)).containsEntry("name", "Alice").containsEntry("dept", "研发");
        assertThat(rows.get(1)).containsEntry("name", "Bob").containsEntry("dept", "测试");
    }

    private static byte[] buildXlsx(String[][] data) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("Sheet1");
            for (int r = 0; r < data.length; r++) {
                Row row = sheet.createRow(r);
                for (int c = 0; c < data[r].length; c++) {
                    row.createCell(c).setCellValue(data[r][c]);
                }
            }
            wb.write(out);
            return out.toByteArray();
        }
    }

    @ExcelInfo(sheetName = "dev", isBigData = false)
    public static class ExportModel {
        @ExcelData
        private List<Dev> data;

        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelColumn(columnName = "部门", index = 2)
        private String dept;

        public ExportModel setData(List<Dev> data) {
            this.data = data;
            return this;
        }

        public List<Dev> getData() {
            return data;
        }
    }

    public static class Dev {
        private final String name;
        private final String dept;

        public Dev(String name, String dept) {
            this.name = name;
            this.dept = dept;
        }

        public String getName() {
            return name;
        }

        public String getDept() {
            return dept;
        }
    }

    @ExcelInfo(sheetName = "dev")
    public static class ImportRow {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelColumn(columnName = "部门", index = 2)
        private String dept;

        public String getName() {
            return name;
        }

        public String getDept() {
            return dept;
        }
    }
}
