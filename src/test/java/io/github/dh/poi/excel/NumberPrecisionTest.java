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
import io.github.dh.poi.excel.importor.StreamingExcelReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/** Large integers must keep full precision (written as text, not a lossy double). */
class NumberPrecisionTest {

    @Test
    void largeLongKeepsFullPrecisionWhileSmallNumbersStayNumeric() throws Exception {
        long bigId = 1234567890123456789L; // 19 digits — not exactly representable as double
        byte[] bytes = ExcelUtil.toBytes(new M().setData(List.of(new Row(bigId, 42))));

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet s = wb.getSheetAt(0);
            Cell idCell = s.getRow(1).getCell(0);
            assertThat(idCell.getCellType()).isEqualTo(CellType.STRING);
            assertThat(idCell.getStringCellValue()).isEqualTo("1234567890123456789");

            Cell qtyCell = s.getRow(1).getCell(1);
            assertThat(qtyCell.getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(qtyCell.getNumericCellValue()).isEqualTo(42.0);
        }
    }

    @Test
    void largeLongSurvivesRoundTripThroughStreamingReader() throws Exception {
        long bigId = 1234567890123456789L;
        byte[] bytes = ExcelUtil.toBytes(new M().setData(List.of(new Row(bigId, 42))));
        List<M> rows = StreamingExcelReader.readAsBeans(new ByteArrayInputStream(bytes), 0, M.class);
        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).id).isEqualTo(bigId);   // exact, not corrupted via double
        assertThat(rows.get(0).qty).isEqualTo(42);
    }

    @ExcelInfo(sheetName = "n", isBigData = false)
    public static class M {
        @ExcelData
        private List<Row> data;

        @ExcelColumn(columnName = "ID", index = 1)
        private Long id;

        @ExcelColumn(columnName = "数量", index = 2)
        private Integer qty;

        public M setData(List<Row> data) { this.data = data; return this; }
        public List<Row> getData() { return data; }
    }

    public static class Row {
        private final Long id;
        private final Integer qty;

        public Row(Long id, Integer qty) {
            this.id = id;
            this.qty = qty;
        }

        public Long getId() { return id; }
        public Integer getQty() { return qty; }
    }
}
