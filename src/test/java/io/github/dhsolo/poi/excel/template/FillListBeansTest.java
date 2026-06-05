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
package io.github.dhsolo.poi.excel.template;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/** Covers ExcelTemplateFiller.fillListBeans (POJO list expansion). */
class FillListBeansTest {

    @Test
    void fillListBeansExpandsRows() throws Exception {
        Workbook template = new XSSFWorkbook();
        Sheet sheet = template.createSheet("Sheet1");
        sheet.createRow(0).createCell(0).setCellValue("标题行");
        Row listRow = sheet.createRow(1);
        listRow.createCell(0).setCellValue("${item.no}");
        listRow.createCell(1).setCellValue("${item.name}");
        sheet.createRow(2).createCell(0).setCellValue("合计");

        List<Item> data = List.of(new Item("1", "设备A"), new Item("2", "设备B"));

        byte[] result = ExcelTemplateFiller.of(template).fillListBeans("item", data).toBytes();

        try (Workbook out = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            Sheet s = out.getSheetAt(0);
            assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("1");
            assertThat(s.getRow(1).getCell(1).getStringCellValue()).isEqualTo("设备A");
            assertThat(s.getRow(2).getCell(0).getStringCellValue()).isEqualTo("2");
            assertThat(s.getRow(2).getCell(1).getStringCellValue()).isEqualTo("设备B");
            // The footer that followed the list row is pushed down below the expanded rows.
            assertThat(s.getRow(3).getCell(0).getStringCellValue()).isEqualTo("合计");
        }
    }

    public static class Item {
        private final String no;
        private final String name;

        public Item(String no, String name) {
            this.no = no;
            this.name = name;
        }

        public String getNo() { return no; }
        public String getName() { return name; }
    }
}
