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
package io.github.dhsolo.poi.excel.core;

import io.github.dhsolo.poi.excel.ExcelModel;
import io.github.dhsolo.poi.excel.ExcelUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * POI-contract conformance of the CSV adapter and the streaming sheet.
 *
 * <p>Regressions guarded: {@code getLastCellNum} used to return the last index instead of
 * index-plus-one (so the trailing required-column check never ran for CSV), out-of-range
 * {@code getRow}/{@code getCell} threw instead of returning null, the row iterators returned
 * null (for-each NPE), and an empty {@code BusinessSXSSFSheet} reported last row 0 instead
 * of -1.
 */
class AdapterContractTest {

    @Test
    void csvLastCellNumFollowsPoiContract() throws Exception {
        CSVSheet sheet = csvSheet("a,b,c\n");
        Row row = sheet.getRow(0);
        assertThat(row.getLastCellNum()).isEqualTo((short) 3); // last index + 1
        assertThat(row.getFirstCellNum()).isEqualTo((short) 0);
        assertThat(row.getCell(2)).isNotNull();
        assertThat(row.getCell(3)).isNull();   // out of range -> null, not IOOBE
        assertThat(row.getCell(-1)).isNull();
    }

    @Test
    void csvGetRowOutOfRangeReturnsNull() throws Exception {
        CSVSheet sheet = csvSheet("a\nb\n");
        assertThat(sheet.getRow(1)).isNotNull();
        assertThat(sheet.getRow(2)).isNull();
        assertThat(sheet.getRow(-1)).isNull();
        assertThat(sheet.getPhysicalNumberOfRows()).isEqualTo(2);
    }

    @Test
    void csvSheetIsIterable() throws Exception {
        CSVSheet sheet = csvSheet("a\nb\nc\n");
        int count = 0;
        for (Row r : sheet) {
            assertThat(r).isNotNull();
            count++;
        }
        assertThat(count).isEqualTo(3);
    }

    /** With the contract fixed, CSV gains the same trailing required-column check as POI sheets. */
    @Test
    void csvMissingTrailingRequiredColumnIsReported() {
        byte[] csv = "名称,级别\nalpha\n".getBytes(StandardCharsets.UTF_8);
        assertThatThrownBy(() -> ExcelUtil.importExcelToMap(new ByteArrayInputStream(csv), 0,
                ExcelModel.of("name"), new ExcelModel("level", false)))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("must not be empty");
    }

    @Test
    void csvImportStillParsesCompleteRows() {
        byte[] csv = "名称,级别\nalpha,高\n".getBytes(StandardCharsets.UTF_8);
        List<Map<String, Object>> rows = ExcelUtil.importExcelToMap(new ByteArrayInputStream(csv), 0,
                ExcelModel.of("name"), ExcelModel.of("level"));
        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).get("name")).isEqualTo("alpha");
        assertThat(rows.get(0).get("level")).isEqualTo("高");
    }

    @Test
    void emptyStreamingSheetReportsMinusOneLastRow() throws Exception {
        try (BusinessSXSSFWorkbook wb = new BusinessSXSSFWorkbook(new BusinessXSSFWorkbook())) {
            Sheet sheet = wb.createSheet("s");
            assertThat(sheet.getLastRowNum()).isEqualTo(-1);
            sheet.createRow(0);
            assertThat(sheet.getLastRowNum()).isEqualTo(0);
            wb.dispose();
        }
    }

    private static CSVSheet csvSheet(String content) throws Exception {
        return new CSVSheet(new ByteArrayInputStream(content.getBytes(StandardCharsets.UTF_8)));
    }
}
