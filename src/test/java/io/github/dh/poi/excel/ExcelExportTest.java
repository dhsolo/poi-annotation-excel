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

import io.github.dh.poi.excel.annotation.ExcelColumn;
import io.github.dh.poi.excel.annotation.ExcelData;
import io.github.dh.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Exercises the core annotation-driven export path end-to-end (covering the getter resolution,
 * inline {@code translate}, the order-number column, and typed cell writing), then reads the
 * workbook back with POI to assert the produced content.
 */
class ExcelExportTest {

    @Test
    void exportsHeaderOrderTranslateAndTypedValues() throws Exception {
        List<Device> data = List.of(
                new Device("路由器", 1, 12.5),
                new Device("交换机", 0, 8.0));

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelUtil.export(out, "devices.xlsx", new DeviceModel().setData(data));

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);

            // Header row: order column "序号" first, then the three columns in index order.
            Row header = sheet.getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("序号");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("名称");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("状态");
            assertThat(header.getCell(3).getStringCellValue()).isEqualTo("数量");

            // First data row.
            Row r1 = sheet.getRow(1);
            assertThat((int) r1.getCell(0).getNumericCellValue()).isEqualTo(1);      // sequence number
            assertThat(r1.getCell(1).getStringCellValue()).isEqualTo("路由器");        // plain getter value
            assertThat(r1.getCell(2).getStringCellValue()).isEqualTo("在线");          // translate 1 -> 在线
            Cell qty = r1.getCell(3);
            assertThat(qty.getCellType()).isEqualTo(CellType.NUMERIC);                // Double written as numeric
            assertThat(qty.getNumericCellValue()).isEqualTo(12.5);

            // Second data row: sequence 2, translate 0 -> 离线.
            Row r2 = sheet.getRow(2);
            assertThat((int) r2.getCell(0).getNumericCellValue()).isEqualTo(2);
            assertThat(r2.getCell(2).getStringCellValue()).isEqualTo("离线");
        }
    }

    @ExcelInfo(sheetName = "设备", needOrder = true, isBigData = false)
    public static class DeviceModel {
        @ExcelData
        private List<Device> data;

        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelColumn(columnName = "状态", index = 2, translate = {"0:离线", "1:在线"})
        private Integer status;

        @ExcelColumn(columnName = "数量", index = 3)
        private Double qty;

        public DeviceModel setData(List<Device> data) {
            this.data = data;
            return this;
        }

        public List<Device> getData() {
            return data;
        }
    }

    public static class Device {
        private final String name;
        private final Integer status;
        private final Double qty;

        public Device(String name, Integer status, Double qty) {
            this.name = name;
            this.status = status;
            this.qty = qty;
        }

        public String getName() {
            return name;
        }

        public Integer getStatus() {
            return status;
        }

        public Double getQty() {
            return qty;
        }
    }
}
