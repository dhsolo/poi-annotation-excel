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
import io.github.dhsolo.poi.excel.annotation.ExcelData;
import io.github.dhsolo.poi.excel.annotation.ExcelFormula;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import io.github.dhsolo.poi.excel.annotation.ExcelListBox;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The validation/formula row coverage used to be hardcoded to 1000 rows; it is now
 * configurable via {@code @ExcelInfo(validateRowCount)} (and
 * {@code ExcelCreator.setValidationRowCount} / the builder).
 */
class ValidationRowCountTest {

    @Test
    void dropdownRangeHonoursConfiguredRowCount() {
        Model m = new Model();
        m.rows = List.of(new Bean("n", "低"));
        try (ExcelCreator creator = new ExcelCreator(m)) {
            creator.createExcel();
            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            List<? extends DataValidation> validations = sheet.getDataValidations();
            assertThat(validations).isNotEmpty();
            CellRangeAddress region = validations.get(0).getRegions().getCellRangeAddress(0);
            // no title: header=0, data starts at 1; 50 rows -> last covered row = 50
            assertThat(region.getFirstRow()).isEqualTo(1);
            assertThat(region.getLastRow()).isEqualTo(50);
        }
    }

    @Test
    void formulaPreFillHonoursConfiguredRowCount() {
        FormulaModel m = new FormulaModel();
        m.rows = List.of(new Bean("n", "低"));
        try (ExcelCreator creator = new ExcelCreator(m)) {
            creator.createExcel();
            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            // formula cells in column 1, rows 1..5 (5 rows from the first data row)
            assertThat(formulaCellCount(sheet, 1)).isEqualTo(5);
        }
    }

    private static int formulaCellCount(Sheet sheet, int column) {
        int count = 0;
        for (Row row : sheet) {
            Cell cell = row.getCell(column);
            if (cell != null && cell.getCellType() == CellType.FORMULA) count++;
        }
        return count;
    }

    @ExcelInfo(sheetName = "v", validateRowCount = 50)
    public static class Model {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelListBox(listTextBox = {"低", "中", "高"})
        @ExcelColumn(columnName = "级别", index = 2)
        private String level;

        @ExcelData
        public List<Bean> rows;
    }

    @ExcelInfo(sheetName = "f", validateRowCount = 5)
    public static class FormulaModel {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelFormula(formula = "LEN(@Column(name))")
        @ExcelColumn(columnName = "长度", index = 2)
        private String len;

        @ExcelData
        public List<Bean> rows;
    }

    public static class Bean {
        private final String name;
        private final String level;

        public Bean(String name, String level) {
            this.name = name;
            this.level = level;
        }

        public String getName() { return name; }
        public String getLevel() { return level; }
    }
}
