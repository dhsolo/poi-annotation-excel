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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: a comparator-based translate map ({@code TreeMap<String,..>}) combined with a
 * non-string cell value (boolean) must not blow up — {@code TreeMap.containsKey} throws
 * {@code ClassCastException} for an incompatible key type, so the lookup falls back to the
 * stringified scan.
 */
class TranslateMapTypeTest {

    @Test
    void treeMapTranslateWithBooleanCellFallsBackToStringScan() throws Exception {
        byte[] xlsx;
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("Sheet1");
            sheet.createRow(0).createCell(0).setCellValue("启用");
            sheet.createRow(1).createCell(0).setCellValue(true); // boolean cell
            wb.write(out);
            xlsx = out.toByteArray();
        }

        Map<Object, Object> translate = new TreeMap<>(); // comparator-based map
        translate.put("true", "是");
        translate.put("false", "否");

        ExcelImportor importor = new ExcelImportor(new ByteArrayInputStream(xlsx));
        LinkedList<ExcelModel> models = new LinkedList<>();
        models.add(ExcelImportor.generateExcelModel("enabled", translate));
        importor.addColumnName(models);
        importor.setStartRow(1);

        assertThat(importor.analysisExcel()).isTrue();
        List<Map> rows = importor.getObject(0, Map.class);
        assertThat(rows.get(0).get("enabled")).isEqualTo("是");
    }
}
