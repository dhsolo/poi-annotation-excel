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
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/** Verifies the CSV read path of ExcelImportor (non-OLE2/OOXML input → CSVWorkbook). */
class CsvImportTest {

    @Test
    void parsesCsvViaImportor() {
        String csv = "name,age\nAlice,30\nBob,25\n";
        ExcelImportor importor = new ExcelImportor(
                new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)));
        importor.addColumnName(columns("name", "age"));
        importor.setStartRow(1);

        boolean ok = importor.analysisExcel();
        assertThat(ok).isTrue();

        @SuppressWarnings("unchecked")
        List<Map<String, Object>> rows = (List<Map<String, Object>>) (List<?>) importor.getObject(0, Map.class);
        assertThat(rows).hasSize(2);
        assertThat(rows.get(0).get("name")).isEqualTo("Alice");
        assertThat(rows.get(0).get("age")).isEqualTo("30");
        assertThat(rows.get(1).get("name")).isEqualTo("Bob");
    }

    private static LinkedList<ExcelModel> columns(String... fields) {
        LinkedList<ExcelModel> list = new LinkedList<>();
        for (String f : fields) list.add(ExcelModel.of(f));
        return list;
    }
}
