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

    @Test
    void importPopulatesNestedChildObject() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new OrderModel().setData(List.of(
                new OrderRow("A001", new Customer("张三", "10086")))));

        // Import the flat sheet back into a model whose @ExcelInfoChild rebuilds the nested object.
        List<OrderImport> rows = ExcelUtil.importExcel(new ByteArrayInputStream(bytes), 0, 1, OrderImport.class);

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getOrderNo()).isEqualTo("A001");
        assertThat(rows.get(0).getCustomer()).isNotNull();
        assertThat(rows.get(0).getCustomer().getName()).isEqualTo("张三");
        assertThat(rows.get(0).getCustomer().getPhone()).isEqualTo("10086");
    }

    @ExcelInfo(sheetName = "order")
    public static class OrderImport {
        @ExcelColumn(columnName = "订单号", index = 1)
        private String orderNo;

        @ExcelInfoChild
        private CustomerImport customer;

        public String getOrderNo() { return orderNo; }
        public void setOrderNo(String orderNo) { this.orderNo = orderNo; }
        public CustomerImport getCustomer() { return customer; }
        public void setCustomer(CustomerImport customer) { this.customer = customer; }
    }

    public static class CustomerImport {
        @ExcelColumn(columnName = "客户名", index = 2)
        private String name;

        @ExcelColumn(columnName = "电话", index = 3)
        private String phone;

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
        public String getPhone() { return phone; }
        public void setPhone(String phone) { this.phone = phone; }
    }

    @Test
    void importDistinguishesChildrenWithSameFieldName() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new TwoChildExport().setData(List.of(
                new TwoChildRow("张三", "李四"))));

        List<TwoChildImport> rows = ExcelUtil.importExcel(new ByteArrayInputStream(bytes), 0, 1, TwoChildImport.class);

        assertThat(rows).hasSize(1);
        // Both children declare a field named "name"; values must not be crossed.
        assertThat(rows.get(0).getCustomer().getName()).isEqualTo("张三");
        assertThat(rows.get(0).getSupplier().getName()).isEqualTo("李四");
    }

    // --- export shapes (two nested objects whose flattened columns share the child field "name") ---
    @ExcelInfo(sheetName = "tc", isBigData = false)
    public static class TwoChildExport {
        @ExcelData private List<TwoChildRow> data;
        @ExcelInfoChild private Named customer;
        @ExcelInfoChild private Named2 supplier;
        public TwoChildExport setData(List<TwoChildRow> d) { this.data = d; return this; }
        public List<TwoChildRow> getData() { return data; }
    }
    public static class Named {
        @ExcelColumn(columnName = "客户名", index = 1) private String name;
        public Named() {} public Named(String n) { this.name = n; }
        public String getName() { return name; } public void setName(String n) { this.name = n; }
    }
    public static class Named2 {
        @ExcelColumn(columnName = "供应商名", index = 2) private String name;
        public Named2() {} public Named2(String n) { this.name = n; }
        public String getName() { return name; } public void setName(String n) { this.name = n; }
    }
    public static class TwoChildRow {
        private final Named customer; private final Named2 supplier;
        public TwoChildRow(String c, String s) { this.customer = new Named(c); this.supplier = new Named2(s); }
        public Named getCustomer() { return customer; } public Named2 getSupplier() { return supplier; }
    }
    // --- import shape ---
    @ExcelInfo(sheetName = "tc")
    public static class TwoChildImport {
        @ExcelInfoChild private Named customer;
        @ExcelInfoChild private Named2 supplier;
        public Named getCustomer() { return customer; } public void setCustomer(Named c) { this.customer = c; }
        public Named2 getSupplier() { return supplier; } public void setSupplier(Named2 s) { this.supplier = s; }
    }

    @Test
    void flattensTwoLevelsAndRoundTrips() throws Exception {
        byte[] bytes = ExcelUtil.toBytes(new L0Export().setData(List.of(
                new L0Row("订单A", new L1("城市X", new L2("100100"))))));

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet s = wb.getSheetAt(0);
            assertThat(s.getRow(0).getCell(0).getStringCellValue()).isEqualTo("单号");
            assertThat(s.getRow(0).getCell(1).getStringCellValue()).isEqualTo("城市");
            assertThat(s.getRow(0).getCell(2).getStringCellValue()).isEqualTo("邮编"); // two levels deep
            assertThat(s.getRow(1).getCell(2).getStringCellValue()).isEqualTo("100100");
        }

        List<L0Import> rows = ExcelUtil.importExcel(new ByteArrayInputStream(bytes), 0, 1, L0Import.class);
        assertThat(rows.get(0).getNo()).isEqualTo("订单A");
        assertThat(rows.get(0).getAddr().getCity()).isEqualTo("城市X");
        assertThat(rows.get(0).getAddr().getGeo().getZip()).isEqualTo("100100"); // nested object rebuilt 2 levels
    }

    // export shapes
    @ExcelInfo(sheetName = "l", isBigData = false)
    public static class L0Export {
        @ExcelData private List<L0Row> data;
        @ExcelColumn(columnName = "单号", index = 1) private String no;
        @ExcelInfoChild private L1 addr;
        public L0Export setData(List<L0Row> d) { this.data = d; return this; }
        public List<L0Row> getData() { return data; }
    }
    public static class L1 {
        @ExcelColumn(columnName = "城市", index = 2) private String city;
        @ExcelInfoChild private L2 geo;
        public L1() {} public L1(String city, L2 geo) { this.city = city; this.geo = geo; }
        public String getCity() { return city; } public void setCity(String c) { this.city = c; }
        public L2 getGeo() { return geo; } public void setGeo(L2 g) { this.geo = g; }
    }
    public static class L2 {
        @ExcelColumn(columnName = "邮编", index = 3) private String zip;
        public L2() {} public L2(String zip) { this.zip = zip; }
        public String getZip() { return zip; } public void setZip(String z) { this.zip = z; }
    }
    public static class L0Row {
        private final String no; private final L1 addr;
        public L0Row(String no, L1 addr) { this.no = no; this.addr = addr; }
        public String getNo() { return no; } public L1 getAddr() { return addr; }
    }
    // import shape
    @ExcelInfo(sheetName = "l")
    public static class L0Import {
        @ExcelColumn(columnName = "单号", index = 1) private String no;
        @ExcelInfoChild private L1 addr;
        public String getNo() { return no; } public void setNo(String n) { this.no = n; }
        public L1 getAddr() { return addr; } public void setAddr(L1 a) { this.addr = a; }
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
