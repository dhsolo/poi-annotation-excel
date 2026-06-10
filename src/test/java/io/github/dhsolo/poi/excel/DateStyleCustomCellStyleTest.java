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
import io.github.dhsolo.poi.excel.annotation.ExcelDateFormat;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;

import java.util.Date;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: date cell styles are cloned from the data style per format pattern and cached;
 * the cache was not invalidated when a custom data style was set afterwards, so date cells
 * kept the stale default look (fill/border/font) while their neighbours used the custom style.
 */
class DateStyleCustomCellStyleTest {

    @Test
    void dateCellsPickUpCustomDataStyle() {
        Model m = new Model();
        m.rows = List.of(new Bean("a", new Date()));
        try (ExcelCreator creator = new ExcelCreator(m)) {
            CellStyle custom = creator.copyCellStyle();
            custom.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            custom.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            creator.setCellStyle(custom);
            creator.createExcel();

            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            Cell dateCell = sheet.getRow(1).getCell(1); // header=0, data=1; date col idx 1
            assertThat(dateCell.getCellStyle().getFillForegroundColor())
                    .isEqualTo(IndexedColors.LIGHT_YELLOW.getIndex());
            // still a date-formatted cell
            assertThat(dateCell.getCellStyle().getDataFormatString()).contains("yyyy");
        }
    }

    @ExcelInfo(sheetName = "d")
    public static class Model {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelDateFormat(pattern = "yyyy-MM-dd")
        @ExcelColumn(columnName = "日期", index = 2)
        private Date day;

        @ExcelData
        public List<Bean> rows;
    }

    public static class Bean {
        private final String name;
        private final Date day;

        public Bean(String name, Date day) {
            this.name = name;
            this.day = day;
        }

        public String getName() { return name; }
        public Date getDay() { return day; }
    }
}
