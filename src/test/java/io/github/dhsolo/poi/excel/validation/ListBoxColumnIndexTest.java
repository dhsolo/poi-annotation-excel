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

    /**
     * Regression: a multi-image picture column expands the physical layout, but realIndex was
     * computed before picture analysis — a dropdown right of the expansion landed on the wrong
     * physical column.
     */
    @Test
    void dropdownShiftsByMultiImageExpansion(@org.junit.jupiter.api.io.TempDir java.io.File tmp) throws Exception {
        java.io.File png = new java.io.File(tmp, "p.png");
        java.awt.image.BufferedImage img = new java.awt.image.BufferedImage(5, 5, java.awt.image.BufferedImage.TYPE_INT_RGB);
        assertThat(javax.imageio.ImageIO.write(img, "png", png)).isTrue();

        PicListBoxModel m = new PicListBoxModel();
        m.rows = List.of(new PicBean(png.getAbsolutePath() + "," + png.getAbsolutePath(), "低"));
        // picture data idx 0 expands to physical cols 0-1; level data idx 1 -> physical col 2
        assertDropdownColumn(m, 2);
    }

    /**
     * Regression: complex-section child creators never ran initHelpers, so their dataValidator
     * stayed null and @ExcelListBox/@ExcelFormula on child sections were silently skipped.
     */
    @Test
    void complexChildSectionDropdownIsApplied() {
        ComplexParent parent = new ComplexParent();
        parent.rows = List.of(new NameBean("x"));
        ChildSection child = new ChildSection();
        child.rows = List.of(new LevelBean("n", "低"));
        parent.child = child;
        try (ExcelCreator creator = new ExcelCreator(parent)) {
            creator.createExcel();
            Sheet sheet = creator.getWorkBook().getSheetAt(0);
            // layout: parent header=0, parent data=1, child header=2, child data starts at 3
            List<? extends DataValidation> validations = sheet.getDataValidations();
            assertThat(validations).isNotEmpty();
            CellRangeAddress region = validations.get(0).getRegions().getCellRangeAddress(0);
            assertThat(region.getFirstRow()).isEqualTo(3);
            assertThat(region.getFirstColumn()).isEqualTo(1);
        }
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

    @ExcelInfo(sheetName = "picList")
    public static class PicListBoxModel {
        @ExcelColumn(columnName = "图片", index = 1)
        @io.github.dhsolo.poi.excel.annotation.ExcelImage
        private String pic;

        @ExcelListBox(listTextBox = {"低", "中", "高"})
        @ExcelColumn(columnName = "级别", index = 2)
        private String level;

        @io.github.dhsolo.poi.excel.annotation.ExcelData
        public List<PicBean> rows;
    }

    public static class PicBean {
        private final String pic;
        private final String level;

        public PicBean(String pic, String level) {
            this.pic = pic;
            this.level = level;
        }

        public String getPic() { return pic; }
        public String getLevel() { return level; }
    }

    @ExcelInfo(sheetName = "complexList")
    public static class ComplexParent implements io.github.dhsolo.poi.excel.model.ComplexExcelModel {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @io.github.dhsolo.poi.excel.annotation.ExcelData
        public List<NameBean> rows;

        ChildSection child;

        @Override
        @SuppressWarnings("rawtypes")
        public List getComplexModels() { return List.of(child); }
    }

    @ExcelInfo(sheetName = "childSection")
    public static class ChildSection {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelListBox(listTextBox = {"低", "中", "高"})
        @ExcelColumn(columnName = "级别", index = 2)
        private String level;

        @io.github.dhsolo.poi.excel.annotation.ExcelData
        public List<LevelBean> rows;
    }

    public static class NameBean {
        private final String name;

        public NameBean(String name) {
            this.name = name;
        }

        public String getName() { return name; }
    }

    public static class LevelBean {
        private final String name;
        private final String level;

        public LevelBean(String name, String level) {
            this.name = name;
            this.level = level;
        }

        public String getName() { return name; }
        public String getLevel() { return level; }
    }
}
