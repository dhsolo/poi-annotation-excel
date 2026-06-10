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
package io.github.dhsolo.poi.excel;

import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import io.github.dhsolo.poi.excel.annotation.ExcelData;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import io.github.dhsolo.poi.excel.annotation.ExcelTitle;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: a sheet with a title and a single column produced a one-cell title "merge",
 * which POI rejects ("Merged region must contain 2 or more cells") and crashed the export.
 */
class SingleColumnTitleTest {

    @Test
    void singleColumnSheetWithTitleExports() {
        Model m = new Model();
        m.rows = List.of(new Bean("a"), new Bean("b"));
        try (ExcelCreator creator = new ExcelCreator(m)) {
            creator.createExcel();
            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("标题");
            assertThat(sheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("a");
            assertThat(sheet.getMergedRegions()).isEmpty();
        }
    }

    @ExcelInfo(sheetName = "single")
    public static class Model {
        @ExcelTitle
        public String title = "标题";

        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelData
        public List<Bean> rows;
    }

    public static class Bean {
        private final String name;

        public Bean(String name) {
            this.name = name;
        }

        public String getName() { return name; }
    }
}
