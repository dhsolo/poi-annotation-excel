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
import io.github.dhsolo.poi.excel.annotation.ExcelTranslateMethod;
import io.github.dhsolo.poi.excel.model.ExcelRowData;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.function.Function;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: the column clones produced for {@code mergeCellIndex > 1} used to be created by
 * Java-serialisation deep cloning, which threw {@code NotSerializableException} as soon as the
 * column carried a non-serialisable member — e.g. the {@code Function} lambda supplied by
 * {@code @ExcelTranslateMethod}. Cloning is now a field-by-field shallow copy.
 */
class MergeCellIndexCloneTest {

    @Test
    void mergeCellIndexColumnWithTranslateMethodExports() {
        Model m = new Model();
        m.rows = List.of(new Bean("a"));
        try (ExcelCreator creator = new ExcelCreator(m)) {
            creator.createExcel();
            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            // header spans two physical columns with the same label
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("名称");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("名称");
            // handler result lands in the data row
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("A!");
        }
    }

    @ExcelInfo(sheetName = "clone")
    public static class Model {
        @ExcelColumn(columnName = "名称", index = 1, mergeCellIndex = 2)
        private String name;

        @ExcelData
        public List<Bean> rows;

        @ExcelTranslateMethod(columnName = "name")
        public Function<ExcelRowData, Object> upper() {
            return d -> String.valueOf(d.currentValue()).toUpperCase() + "!";
        }
    }

    public static class Bean {
        private final String name;

        public Bean(String name) {
            this.name = name;
        }

        public String getName() { return name; }
    }
}
