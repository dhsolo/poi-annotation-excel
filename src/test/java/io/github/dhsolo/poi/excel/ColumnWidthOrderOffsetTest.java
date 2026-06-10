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
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: {@code @ExcelColumn(columnWidth)} is keyed by data column but used to be applied
 * to the physical column of the same index, so with {@code needOrder = true} every custom width
 * landed one column to the left (on the order column).
 */
class ColumnWidthOrderOffsetTest {

    private static final int DEFAULT_WIDTH = 20 * 255;

    @Test
    void annotationWidthLandsOnItsPhysicalColumnWithOrder() {
        Model m = new Model();
        m.rows = List.of(new Bean("a"));
        try (ExcelCreator creator = new ExcelCreator(m)) {
            creator.createExcel();
            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            // physical layout: col0 = order, col1 = 名称(columnWidth=30)
            assertThat(sheet.getColumnWidth(0)).isEqualTo(DEFAULT_WIDTH);
            assertThat(sheet.getColumnWidth(1)).isEqualTo(30 * 255);
        }
    }

    @ExcelInfo(sheetName = "w", needOrder = true)
    public static class Model {
        @ExcelColumn(columnName = "名称", index = 1, columnWidth = 30)
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
