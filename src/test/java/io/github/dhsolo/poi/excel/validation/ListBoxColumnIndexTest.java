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
package io.github.dhsolo.poi.excel.validation;

import io.github.dhsolo.poi.excel.ExcelCreator;
import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import io.github.dhsolo.poi.excel.annotation.ExcelListBox;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: the dropdown-validation column ({@code realIndex}) must shift by the whole order
 * block. With {@code orderColumnSpan > 1} it used to shift by one only, putting the dropdown on
 * the wrong physical column ({@code orderColumnSpan} was also assigned after the index
 * recomputation it feeds, so the span never took effect).
 */
class ListBoxColumnIndexTest {

    @Test
    void dropdownLandsOnItsColumnWithPlainOrder() {
        assertDropdownColumn(new OrderedModel(), 2); // order=col0, name=col1, level=col2
    }

    @Test
    void dropdownShiftsByWholeOrderBlock() {
        assertDropdownColumn(new SpanModel(), 3); // order block=cols0-1, name=col2, level=col3
    }

    private static void assertDropdownColumn(Object model, int expectedColumn) {
        try (ExcelCreator creator = new ExcelCreator(model)) {
            creator.createExcel();
            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            List<? extends DataValidation> validations = sheet.getDataValidations();
            assertThat(validations).isNotEmpty();
            CellRangeAddress region = validations.get(0).getRegions().getCellRangeAddress(0);
            assertThat(region.getFirstColumn()).isEqualTo(expectedColumn);
            assertThat(region.getLastColumn()).isEqualTo(expectedColumn);
        }
    }

    @ExcelInfo(sheetName = "ordered", needOrder = true)
    public static class OrderedModel {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelListBox(listTextBox = {"低", "中", "高"})
        @ExcelColumn(columnName = "级别", index = 2)
        private String level;
    }

    @ExcelInfo(sheetName = "span", needOrder = true, orderColumnSpan = 2)
    public static class SpanModel {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelListBox(listTextBox = {"低", "中", "高"})
        @ExcelColumn(columnName = "级别", index = 2)
        private String level;
    }
}
