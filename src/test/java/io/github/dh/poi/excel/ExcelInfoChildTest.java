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
import io.github.dh.poi.excel.annotation.ExcelInfoChild;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/** Verifies @ExcelInfoChild flattens a nested object's columns into the parent sheet (ordered by index). */
class ExcelInfoChildTest {

    @Test
    void flattensNestedChildColumns() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new OrderModel().setData(List.of(
                new OrderRow("A001", new Customer("张三", "10086")))));

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet s = wb.getSheetAt(0);
            Row header = s.getRow(0);
            // Columns ordered by @ExcelColumn.index across parent (1) and flattened child (2,3).
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("订单号");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("客户名");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("电话");

            Row r1 = s.getRow(1);
            assertThat(r1.getCell(0).getStringCellValue()).isEqualTo("A001");
            assertThat(r1.getCell(1).getStringCellValue()).isEqualTo("张三");   // customer.name
            assertThat(r1.getCell(2).getStringCellValue()).isEqualTo("10086"); // customer.phone
        }
    }

    @ExcelInfo(sheetName = "order", isBigData = false)
    public static class OrderModel {
        @ExcelData
        private List<OrderRow> data;

        @ExcelColumn(columnName = "订单号", index = 1)
        private String orderNo;

        @ExcelInfoChild
        private Customer customer;   // its @ExcelColumn fields are flattened into this sheet

        public OrderModel setData(List<OrderRow> data) { this.data = data; return this; }
        public List<OrderRow> getData() { return data; }
    }

    public static class Customer {
        @ExcelColumn(columnName = "客户名", index = 2)
        private String name;

        @ExcelColumn(columnName = "电话", index = 3)
        private String phone;

        public Customer() {}
        public Customer(String name, String phone) { this.name = name; this.phone = phone; }
        public String getName() { return name; }
        public String getPhone() { return phone; }
    }

    public static class OrderRow {
        private final String orderNo;
        private final Customer customer;

        public OrderRow(String orderNo, Customer customer) {
            this.orderNo = orderNo;
            this.customer = customer;
        }

        public String getOrderNo() { return orderNo; }
        public Customer getCustomer() { return customer; }
    }
}
