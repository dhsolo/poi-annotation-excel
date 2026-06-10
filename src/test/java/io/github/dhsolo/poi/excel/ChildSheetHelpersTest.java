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
import io.github.dhsolo.poi.excel.annotation.ExcelImage;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import io.github.dhsolo.poi.excel.annotation.ExcelListBox;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: appended child sheets ({@code getChild().add(...)}) built their picture handler
 * and data validator in their own constructor, bound to the child's ORIGINAL (later discarded)
 * workbook/sheet — so pictures and dropdown validations on child sheets landed in an orphaned
 * book and silently vanished from the export. The child now shares the parent's picture handler
 * (rebound to the child's sheet) and rebuilds its validator against the shared workbook.
 */
class ChildSheetHelpersTest {

    @Test
    void childSheetDropdownValidationLandsOnChildSheet() {
        try (ExcelCreator parent = new ExcelCreator(parentModel());
             ExcelCreator child = new ExcelCreator(listBoxModel())) {
            parent.getChild().add(child);
            parent.createExcel();

            org.apache.poi.ss.usermodel.Sheet childSheet = parent.getWorkBook().getSheet("childList");
            assertThat(childSheet).isNotNull();
            List<? extends DataValidation> validations = childSheet.getDataValidations();
            assertThat(validations).isNotEmpty();
            // level column: data idx 1 -> physical col 1 (no order column)
            assertThat(validations.get(0).getRegions().getCellRangeAddress(0).getFirstColumn()).isEqualTo(1);
        }
    }

    @Test
    void childSheetPictureIsEmbeddedInTheSharedWorkbook(@TempDir File tmp) throws Exception {
        File jpg = new File(tmp, "p.jpg");
        BufferedImage img = new BufferedImage(8, 8, BufferedImage.TYPE_INT_RGB);
        assertThat(ImageIO.write(img, "jpg", jpg)).isTrue();

        File out = new File(tmp, "multi.xlsx");
        try (ExcelCreator parent = new ExcelCreator(parentModel());
             ExcelCreator child = new ExcelCreator(pictureModel(jpg.getAbsolutePath()))) {
            parent.getChild().add(child);
            parent.createExcel();
            parent.exportLocal(out.getAbsolutePath());
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(out))) {
            assertThat(wb.getAllPictures()).hasSize(1);
            XSSFSheet childSheet = wb.getSheet("childPics");
            assertThat(childSheet).isNotNull();
            // the anchor must live on the child sheet, not on the parent or an orphaned book
            assertThat(childSheet.getDrawingPatriarch()).isNotNull();
            assertThat(childSheet.getDrawingPatriarch().getShapes()).hasSize(1);
        }
    }

    private static ParentModel parentModel() {
        ParentModel m = new ParentModel();
        m.rows = List.of(new NameBean("x"));
        return m;
    }

    private static ListBoxModel listBoxModel() {
        ListBoxModel m = new ListBoxModel();
        m.rows = List.of(new LevelBean("n", "低"));
        return m;
    }

    private static PictureModel pictureModel(String path) {
        PictureModel m = new PictureModel();
        m.rows = List.of(new PicBean("n", path));
        return m;
    }

    @ExcelInfo(sheetName = "parent")
    public static class ParentModel {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelData
        public List<NameBean> rows;
    }

    @ExcelInfo(sheetName = "childList")
    public static class ListBoxModel {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelListBox(listTextBox = {"低", "中", "高"})
        @ExcelColumn(columnName = "级别", index = 2)
        private String level;

        @ExcelData
        public List<LevelBean> rows;
    }

    @ExcelInfo(sheetName = "childPics")
    public static class PictureModel {
        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelColumn(columnName = "图片", index = 2)
        @ExcelImage
        private String pic;

        @ExcelData
        public List<PicBean> rows;
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

    public static class PicBean {
        private final String name;
        private final String pic;

        public PicBean(String name, String pic) {
            this.name = name;
            this.pic = pic;
        }

        public String getName() { return name; }
        public String getPic() { return pic; }
    }
}
