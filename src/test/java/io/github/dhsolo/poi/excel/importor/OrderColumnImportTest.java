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
import io.github.dhsolo.poi.excel.annotation.ExcelData;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * When a class declares {@link ExcelInfo#needOrder()}, the export prepends an auto-sequence column
 * that is not an {@code @ExcelColumn} field. A class-based import must skip that leading column so
 * the data columns map to the models — otherwise every column is read one position to the left
 * (the sequence number lands in the first field).
 */
class OrderColumnImportTest {

    @Test
    void classImportSkipsLeadingOrderColumn() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new OrderedModel().setData(List.of(
                new OrderedModel("张三", 30),
                new OrderedModel("李四", 25))));

        List<OrderedModel> rows = ExcelUtil.importExcel(new ByteArrayInputStream(bytes), OrderedModel.class);

        assertThat(rows).hasSize(2);
        assertThat(rows.get(0).getName()).isEqualTo("张三");
        assertThat(rows.get(0).getAge()).isEqualTo(30);
        assertThat(rows.get(1).getName()).isEqualTo("李四");
        assertThat(rows.get(1).getAge()).isEqualTo(25);
    }

    @Test
    void orderColumnSpanGreaterThanOneSkipsAllLeadingColumns() throws Exception {
        // Build a file laid out as a span-2 order column would produce: two leading columns before
        // the data columns. The class declares orderColumnSpan = 2, so both must be skipped.
        byte[] bytes;
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("s");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("序号");
            header.createCell(1).setCellValue("序号");
            header.createCell(2).setCellValue("名称");
            header.createCell(3).setCellValue("年龄");
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue(1);
            r1.createCell(1).setCellValue(1);
            r1.createCell(2).setCellValue("王五");
            r1.createCell(3).setCellValue(40);
            wb.write(out);
            bytes = out.toByteArray();
        }

        List<SpanTwoModel> rows = ExcelUtil.importExcel(new ByteArrayInputStream(bytes), SpanTwoModel.class);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getName()).isEqualTo("王五");
        assertThat(rows.get(0).getAge()).isEqualTo(40);
    }

    @ExcelInfo(sheetName = "s", isBigData = false, needOrder = true)
    public static class OrderedModel {
        @ExcelData
        private List<OrderedModel> data;
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;
        @ExcelColumn(columnName = "年龄", index = 2)
        private Integer age;

        public OrderedModel() {
        }

        public OrderedModel(String name, Integer age) {
            this.name = name;
            this.age = age;
        }

        public OrderedModel setData(List<OrderedModel> data) {
            this.data = data;
            return this;
        }

        public List<OrderedModel> getData() {
            return data;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Integer getAge() {
            return age;
        }

        public void setAge(Integer age) {
            this.age = age;
        }
    }

    @ExcelInfo(sheetName = "s", isBigData = false, needOrder = true, orderColumnSpan = 2)
    public static class SpanTwoModel {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;
        @ExcelColumn(columnName = "年龄", index = 2)
        private Integer age;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Integer getAge() {
            return age;
        }

        public void setAge(Integer age) {
            this.age = age;
        }
    }
}
