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
package io.github.dhsolo.poi.excel.picture;

import io.github.dhsolo.poi.excel.ExcelUtil;
import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import io.github.dhsolo.poi.excel.annotation.ExcelData;
import io.github.dhsolo.poi.excel.annotation.ExcelImage;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.apache.poi.ss.usermodel.Row;
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
 * Regression: a cell holding multiple images expands the picture column by extra physical
 * columns; every later column must still map to its own data column. The per-column expansion
 * offset used to be overwritten instead of accumulated, which crashed with an NPE (or shifted
 * values) as soon as two or more ordinary columns followed a multi-image column.
 */
class MultiImagePictureColumnTest {

    @Test
    void columnsAfterMultiImageExpansionKeepTheirValues(@TempDir File tmp) throws Exception {
        String twoPics = twoPicturePaths(tmp);
        List<RowBean> rows = List.of(new RowBean(twoPics, "alpha", "beta"));

        File out = new File(tmp, "out.xlsx");
        ExcelUtil.exportLocal(out.getAbsolutePath(), new Model().setData(rows));

        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(out))) {
            Row data = wb.getSheetAt(0).getRow(1); // row 0 = header (no title)
            // physical layout: col 0-1 picture (expanded), col 2 = colA, col 3 = colB
            assertThat(data.getCell(2).getStringCellValue()).isEqualTo("alpha");
            assertThat(data.getCell(3).getStringCellValue()).isEqualTo("beta");
        }
    }

    @Test
    void multiImageExpansionAlsoWorksWithOrderColumn(@TempDir File tmp) throws Exception {
        String twoPics = twoPicturePaths(tmp);
        List<RowBean> rows = List.of(new RowBean(twoPics, "alpha", "beta"));

        File out = new File(tmp, "out.xlsx");
        ExcelUtil.exportLocal(out.getAbsolutePath(), new OrderedModel().setData(rows));

        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(out))) {
            Row data = wb.getSheetAt(0).getRow(1);
            // physical layout: col 0 = order, col 1-2 picture (expanded), col 3 = colA, col 4 = colB
            assertThat(data.getCell(0).getNumericCellValue()).isEqualTo(1.0);
            assertThat(data.getCell(3).getStringCellValue()).isEqualTo("alpha");
            assertThat(data.getCell(4).getStringCellValue()).isEqualTo("beta");
        }
    }

    /**
     * Regression (complex multi-section sheets): the picture handler is shared across sections
     * and its cross-section expansion mapping mixes colliding data-column keys. A child
     * section's vertical merge / layout must only honour its OWN expansion entries — a parent
     * section with a multi-image column used to shift the child's merge column.
     */
    @Test
    void parentSectionExpansionDoesNotShiftChildSectionMerge(@TempDir File tmp) throws Exception {
        ComplexParent parent = new ComplexParent();
        parent.rows = List.of(new RowBean(twoPicturePaths(tmp), "alpha", "beta"));
        ChildSection child = new ChildSection();
        child.rows = List.of(new MergeBean("A"), new MergeBean("A"), new MergeBean("B"));
        parent.child = child;

        File out = new File(tmp, "complex.xlsx");
        ExcelUtil.exportLocal(out.getAbsolutePath(), parent);

        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(out))) {
            // layout: parent header=0, parent data=1, child header=2, child data=3..5
            // child merge column: data idx 1 -> physical col 1; the parent's multi-image
            // expansion entry (data idx 0, +1 col) must NOT shift it to col 2.
            org.apache.poi.ss.util.CellRangeAddress merge = null;
            for (org.apache.poi.ss.util.CellRangeAddress r : wb.getSheetAt(0).getMergedRegions()) {
                if (r.getFirstRow() != r.getLastRow() && r.getFirstRow() >= 3) {
                    merge = r;
                }
            }
            assertThat(merge).isNotNull();
            assertThat(merge.getFirstColumn()).isEqualTo(1);
            assertThat(merge.getLastColumn()).isEqualTo(1);
            assertThat(merge.getFirstRow()).isEqualTo(3);
            assertThat(merge.getLastRow()).isEqualTo(4);
        }
    }

    @ExcelInfo(sheetName = "complexPics")
    public static class ComplexParent implements io.github.dhsolo.poi.excel.model.ComplexExcelModel {
        @ExcelData
        private List<RowBean> rows;

        @ExcelColumn(columnName = "图片", index = 1)
        @ExcelImage
        private String pic;

        @ExcelColumn(columnName = "列A", index = 2)
        private String colA;

        @ExcelColumn(columnName = "列B", index = 3)
        private String colB;

        ChildSection child;

        public List<RowBean> getRows() { return rows; }

        @Override
        @SuppressWarnings("rawtypes")
        public List getComplexModels() { return List.of(child); }
    }

    @ExcelInfo(sheetName = "childSection")
    public static class ChildSection {
        @ExcelColumn(columnName = "数量", index = 1, sourceField = "group")
        private String count;

        // merge column at data idx 1: only a column RIGHT of the parent's picture key (0)
        // exercises the cross-section shift bug
        @ExcelColumn(columnName = "分组", index = 2, needMergeCell = true)
        private String group;

        @ExcelData
        public List<MergeBean> rows;
    }

    public static class MergeBean {
        private final String group;

        public MergeBean(String group) {
            this.group = group;
        }

        public String getGroup() { return group; }
    }

    /** Writes one tiny PNG and returns "path,path" so a single cell carries two images. */
    private static String twoPicturePaths(File tmp) throws Exception {
        File png = new File(tmp, "p.png");
        BufferedImage img = new BufferedImage(5, 5, BufferedImage.TYPE_INT_RGB);
        assertThat(ImageIO.write(img, "png", png)).isTrue();
        return png.getAbsolutePath() + "," + png.getAbsolutePath();
    }

    @ExcelInfo(sheetName = "pics")
    public static class Model {
        @ExcelData
        private List<RowBean> data;

        @ExcelColumn(columnName = "图片", index = 1)
        @ExcelImage
        private String pic;

        @ExcelColumn(columnName = "列A", index = 2)
        private String colA;

        @ExcelColumn(columnName = "列B", index = 3)
        private String colB;

        public Model setData(List<RowBean> data) { this.data = data; return this; }
        public List<RowBean> getData() { return data; }
    }

    @ExcelInfo(sheetName = "pics", needOrder = true)
    public static class OrderedModel {
        @ExcelData
        private List<RowBean> data;

        @ExcelColumn(columnName = "图片", index = 1)
        @ExcelImage
        private String pic;

        @ExcelColumn(columnName = "列A", index = 2)
        private String colA;

        @ExcelColumn(columnName = "列B", index = 3)
        private String colB;

        public OrderedModel setData(List<RowBean> data) { this.data = data; return this; }
        public List<RowBean> getData() { return data; }
    }

    public static class RowBean {
        private final String pic;
        private final String colA;
        private final String colB;

        public RowBean(String pic, String colA, String colB) {
            this.pic = pic;
            this.colA = colA;
            this.colB = colB;
        }

        public String getPic() { return pic; }
        public String getColA() { return colA; }
        public String getColB() { return colB; }
    }
}
