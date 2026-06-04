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
package io.github.dh.poi.excel.importor;

import io.github.dh.poi.excel.ExcelModel;
import io.github.dh.poi.excel.ExcelUtil;
import io.github.dh.poi.excel.annotation.ExcelColumn;
import io.github.dh.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class ExcelImportorTest {

    /** Builds a minimal XLSX workbook in memory and returns its bytes. */
    private byte[] buildXlsx(String[][] data) throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        for (int r = 0; r < data.length; r++) {
            Row row = sheet.createRow(r);
            for (int c = 0; c < data[r].length; c++) {
                row.createCell(c).setCellValue(data[r][c]);
            }
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        wb.write(out);
        wb.close();
        return out.toByteArray();
    }

    private LinkedList<ExcelModel> columns(String... fields) {
        LinkedList<ExcelModel> list = new LinkedList<>();
        for (String f : fields) list.add(new ExcelModel(f));
        return list;
    }

    @Test
    void parseXlsx_simpleData_returnsRows() throws Exception {
        byte[] xlsx = buildXlsx(new String[][] {
                {"姓名", "年龄"},  // header row (row 0 — skipped by startRow=1)
                {"Alice", "30"},
                {"Bob", "25"},
        });

        ExcelImportor importor = new ExcelImportor(new ByteArrayInputStream(xlsx));
        importor.addColumnName(columns("name", "age"));
        importor.setStartRow(1);
        boolean ok = importor.analysisExcel();

        assertThat(ok).isTrue();
        @SuppressWarnings("unchecked")
        List<Map<String, Object>> rows = (List<Map<String, Object>>) (List<?>) importor.getObject(0, Map.class);
        assertThat(rows).hasSize(2);
        assertThat(rows.get(0).get("name")).isEqualTo("Alice");
        assertThat(rows.get(1).get("name")).isEqualTo("Bob");
    }

    @Test
    void parseXlsx_withListener_receivesRows() throws Exception {
        byte[] xlsx = buildXlsx(new String[][] {
                {"header"},
                {"row1"},
                {"row2"},
        });

        List<Map<String, Object>> received = new ArrayList<>();
        AtomicInteger finishCount = new AtomicInteger(0);

        ExcelImportor importor = new ExcelImportor(new ByteArrayInputStream(xlsx));
        importor.addColumnName(columns("value"));
        importor.setStartRow(1);
        importor.setReadListener(new ExcelReadListener() {
            @Override
            public void onRow(Map<String, Object> row, int rowIndex) {
                received.add(new HashMap<>(row));
            }
            @Override
            public void onFinish(int totalRows) {
                finishCount.set(totalRows);
            }
        });

        importor.analysisExcel();

        assertThat(received).hasSize(2);
        assertThat(received.get(0).get("value")).isEqualTo("row1");
        assertThat(received.get(1).get("value")).isEqualTo("row2");
    }

    @Test
    void getWorkBookType_xlsx_returnsXlsx() throws Exception {
        byte[] xlsx = buildXlsx(new String[][] {{"x"}});
        ExcelImportor importor = new ExcelImportor(new ByteArrayInputStream(xlsx));
        assertThat(importor.getWorkBookType()).isEqualTo("xlsx");
    }

    @ExcelInfo(sheetName = "Sheet1")
    public static class UserModel {
        @ExcelColumn(columnName = "Name", index = 1)
        private String name;
        @ExcelColumn(columnName = "Age", index = 2)
        private Integer age;

        public UserModel() {}
        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
        public Integer getAge() { return age; }
        public void setAge(Integer age) { this.age = age; }
    }

    public static class PlainPojo {
        private String name;
        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
    }

    @Test
    void importExcel_annotatedClass_derivesColumnsAutomatically() throws Exception {
        byte[] xlsx = buildXlsx(new String[][] {
                {"Name", "Age"},
                {"Alice", "30"},
                {"Bob", "25"},
        });

        List<UserModel> users = ExcelUtil.importExcel(new ByteArrayInputStream(xlsx), UserModel.class);

        assertThat(users).hasSize(2);
        assertThat(users.get(0).getName()).isEqualTo("Alice");
        assertThat(users.get(0).getAge()).isEqualTo(30);
        assertThat(users.get(1).getName()).isEqualTo("Bob");
        assertThat(users.get(1).getAge()).isEqualTo(25);
    }

    @Test
    void importExcel_unannotatedClassNoColumns_throwsClearError() throws Exception {
        byte[] xlsx = buildXlsx(new String[][] {{"value"}, {"x"}});

        assertThatThrownBy(() -> ExcelUtil.importExcel(new ByteArrayInputStream(xlsx), PlainPojo.class))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("@ExcelInfo");
    }

    @Test
    void nullableColumn_emptyCell_passesValidation() throws Exception {
        byte[] xlsx = buildXlsx(new String[][] {
                {"value"},
                {""},
        });

        ExcelModel model = new ExcelModel("value", true);
        LinkedList<ExcelModel> cols = new LinkedList<>();
        cols.add(model);

        ExcelImportor importor = new ExcelImportor(new ByteArrayInputStream(xlsx));
        importor.addColumnName(cols);
        importor.setStartRow(1);

        boolean ok = importor.analysisExcel();
        assertThat(ok).isTrue();
    }
}
