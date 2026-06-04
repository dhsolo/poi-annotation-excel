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
package io.github.dh.poi.excel.template;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class ExcelTemplateFillerTest {

    private Workbook buildTemplate(String... cellValues) throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        Row row = sheet.createRow(0);
        for (int i = 0; i < cellValues.length; i++) {
            row.createCell(i).setCellValue(cellValues[i]);
        }
        return wb;
    }

    @Test
    void fill_scalarPlaceholder_replacesValue() throws Exception {
        Workbook template = buildTemplate("日期:", "${date}", "状态:", "${status}");

        byte[] result = ExcelTemplateFiller.of(template)
                .fill("date", "2024-06-01")
                .fill("status", "正常")
                .toBytes();

        Workbook out = new XSSFWorkbook(new ByteArrayInputStream(result));
        Row row = out.getSheetAt(0).getRow(0);
        assertThat(row.getCell(1).getStringCellValue()).isEqualTo("2024-06-01");
        assertThat(row.getCell(3).getStringCellValue()).isEqualTo("正常");
        out.close();
    }

    @Test
    void fill_missingKey_replacesWithEmpty() throws Exception {
        Workbook template = buildTemplate("${missing}");

        byte[] result = ExcelTemplateFiller.of(template).toBytes();

        Workbook out = new XSSFWorkbook(new ByteArrayInputStream(result));
        assertThat(out.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEmpty();
        out.close();
    }

    @Test
    void fillAll_map_replacesAll() throws Exception {
        Workbook template = buildTemplate("${a}", "${b}");

        byte[] result = ExcelTemplateFiller.of(template)
                .fillAll(Map.of("a", "alpha", "b", "beta"))
                .toBytes();

        Workbook out = new XSSFWorkbook(new ByteArrayInputStream(result));
        Row row = out.getSheetAt(0).getRow(0);
        assertThat(row.getCell(0).getStringCellValue()).isEqualTo("alpha");
        assertThat(row.getCell(1).getStringCellValue()).isEqualTo("beta");
        out.close();
    }

    @Test
    void fillList_expandsRows() throws Exception {
        Workbook template = new XSSFWorkbook();
        Sheet sheet = template.createSheet("Sheet1");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("标题行");

        Row listRow = sheet.createRow(1);
        listRow.createCell(0).setCellValue("${item.no}");
        listRow.createCell(1).setCellValue("${item.name}");

        Row footer = sheet.createRow(2);
        footer.createCell(0).setCellValue("合计");

        List<Map<String, Object>> data = List.of(
                Map.of("no", "1", "name", "设备A"),
                Map.of("no", "2", "name", "设备B"),
                Map.of("no", "3", "name", "设备C")
        );

        byte[] result = ExcelTemplateFiller.of(template)
                .fillList("item", data)
                .toBytes();

        Workbook out = new XSSFWorkbook(new ByteArrayInputStream(result));
        Sheet outSheet = out.getSheetAt(0);
        assertThat(outSheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("1");
        assertThat(outSheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("2");
        assertThat(outSheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("3");
        out.close();
    }
}
