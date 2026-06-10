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
import io.github.dhsolo.poi.excel.annotation.ExcelColumnParent;
import io.github.dhsolo.poi.excel.annotation.ExcelData;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import io.github.dhsolo.poi.excel.annotation.ExcelTitle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: the vertical merge driven by {@code needMergeCell} must target the data rows.
 * The merge base row used to ignore the extra grouped-header row added by
 * {@code @ExcelColumnParent}, landing the merge one row too high and crashing with an
 * overlapping-region {@code IllegalStateException}.
 */
class VerticalMergeRowOffsetTest {

    @Test
    void mergeTargetsDataRowsWithoutParentHeader() {
        PlainModel m = new PlainModel();
        m.rows = rows();
        try (ExcelCreator creator = new ExcelCreator(m)) {
            creator.createExcel();
            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            // layout: title=0, header=1, data=2..4 ; values A,A,B -> merge rows 2-3 in col 0
            CellRangeAddress vertical = findVerticalMerge(sheet);
            assertThat(vertical).isNotNull();
            assertThat(vertical.getFirstRow()).isEqualTo(2);
            assertThat(vertical.getLastRow()).isEqualTo(3);
        }
    }

    @Test
    void mergeTargetsDataRowsWithParentHeader() {
        ParentModel m = new ParentModel();
        m.rows = rows();
        try (ExcelCreator creator = new ExcelCreator(m)) {
            creator.createExcel();
            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            // layout: title=0, parent header=1, header=2, data=3..5 ; merge rows 3-4 in col 0
            CellRangeAddress vertical = findDataRowVerticalMerge(sheet, 3);
            assertThat(vertical).isNotNull();
            assertThat(vertical.getFirstRow()).isEqualTo(3);
            assertThat(vertical.getLastRow()).isEqualTo(4);
        }
    }

    /** First multi-row merge in column 0 (no parent header: the only one is the data merge). */
    private static CellRangeAddress findVerticalMerge(Sheet sheet) {
        return findDataRowVerticalMerge(sheet, 0);
    }

    /** Multi-row merge in column 0 starting at or below {@code minFirstRow} (skips header merges). */
    private static CellRangeAddress findDataRowVerticalMerge(Sheet sheet, int minFirstRow) {
        for (CellRangeAddress r : sheet.getMergedRegions()) {
            if (r.getFirstColumn() == 0 && r.getLastColumn() == 0
                    && r.getFirstRow() != r.getLastRow() && r.getFirstRow() >= minFirstRow) {
                return r;
            }
        }
        return null;
    }

    private static List<RowBean> rows() {
        List<RowBean> rows = new ArrayList<>();
        rows.add(new RowBean("A", 1, 2));
        rows.add(new RowBean("A", 3, 4));
        rows.add(new RowBean("B", 5, 6));
        return rows;
    }

    public static class RowBean {
        private final String group;
        private final Integer q1;
        private final Integer q2;

        public RowBean(String group, Integer q1, Integer q2) {
            this.group = group;
            this.q1 = q1;
            this.q2 = q2;
        }

        public String getGroup() { return group; }
        public Integer getQ1() { return q1; }
        public Integer getQ2() { return q2; }
    }

    @ExcelInfo(sheetName = "plain")
    public static class PlainModel {
        @ExcelTitle
        public String title = "标题";

        @ExcelColumn(columnName = "分组", index = 1, needMergeCell = true)
        private String group;

        @ExcelColumn(columnName = "一季度", index = 2, sourceField = "q1")
        private Integer q1;

        @ExcelData
        public List<RowBean> rows;
    }

    @ExcelInfo(sheetName = "parent")
    public static class ParentModel {
        @ExcelTitle
        public String title = "标题";

        @ExcelColumn(columnName = "分组", index = 1, needMergeCell = true)
        private String group;

        @ExcelColumnParent(value = "销售额", columns = {
                @ExcelColumn(columnName = "一季度", index = 2, sourceField = "q1"),
                @ExcelColumn(columnName = "二季度", index = 3, sourceField = "q2")
        })
        private Object sales;

        @ExcelData
        public List<RowBean> rows;
    }
}
