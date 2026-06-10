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
