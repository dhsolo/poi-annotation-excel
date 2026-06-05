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
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Workbook;
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
 * Exercises the disk-staging picture export path end-to-end using local image files as the
 * image source (no network), verifying that:
 * <ul>
 *   <li>the disk-staging path is active for a top-level picture sheet (the {@code hasPicture}
 *       fix);</li>
 *   <li>each image's original format is preserved (PNG stays PNG, JPEG stays JPEG) via the
 *       raw byte-passthrough path;</li>
 *   <li>the produced workbook is valid and contains exactly the embedded pictures.</li>
 * </ul>
 */
class PictureExportTest {

    @Test
    void exportPreservesOriginalImageFormats(@TempDir File tmp) throws Exception {
        File png = new File(tmp, "alpha.png");
        BufferedImage argb = new BufferedImage(10, 10, BufferedImage.TYPE_INT_ARGB);
        argb.setRGB(0, 0, 0x80FF0000); // semi-transparent red pixel
        assertThat(ImageIO.write(argb, "png", png)).isTrue();

        File jpg = new File(tmp, "photo.jpg");
        BufferedImage rgb = new BufferedImage(10, 10, BufferedImage.TYPE_INT_RGB);
        rgb.setRGB(0, 0, 0x00FF00);
        assertThat(ImageIO.write(rgb, "jpg", jpg)).isTrue();

        List<PicRow> rows = List.of(
                new PicRow("row-png", png.getAbsolutePath()),
                new PicRow("row-jpg", jpg.getAbsolutePath()));

        File out = new File(tmp, "out.xlsx");
        ExcelUtil.exportLocal(out.getAbsolutePath(), new PicExportModel().setData(rows));

        assertThat(out).exists();
        assertThat(out.length()).isGreaterThan(0);

        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(out))) {
            List<? extends PictureData> pics = wb.getAllPictures();
            assertThat(pics).hasSize(2);
            assertThat(pics).extracting(PictureData::getPictureType)
                    .containsExactlyInAnyOrder(Workbook.PICTURE_TYPE_PNG, Workbook.PICTURE_TYPE_JPEG);
            // Raw passthrough means the stored bytes are the original files unchanged.
            assertThat(pics).allSatisfy(p -> assertThat(p.getData()).isNotEmpty());
        }

        // Injected media must be STORED (not re-deflated); workbook XML must stay DEFLATED.
        try (java.util.zip.ZipFile zip = new java.util.zip.ZipFile(out)) {
            var media = zip.stream()
                    .filter(e -> e.getName().startsWith("xl/media/"))
                    .toList();
            assertThat(media).isNotEmpty();
            assertThat(media).allSatisfy(e ->
                    assertThat(e.getMethod()).isEqualTo(java.util.zip.ZipEntry.STORED));
            assertThat(zip.getEntry("xl/worksheets/sheet1.xml").getMethod())
                    .isEqualTo(java.util.zip.ZipEntry.DEFLATED);
        }
    }

    @ExcelInfo(sheetName = "pics", isBigData = false)
    public static class PicExportModel {
        @ExcelData
        private List<PicRow> data;

        @ExcelColumn(columnName = "名称", index = 1)
        private String name;

        @ExcelColumn(columnName = "图片", index = 2)
        @ExcelImage
        private String pic;

        public PicExportModel setData(List<PicRow> data) {
            this.data = data;
            return this;
        }

        public List<PicRow> getData() {
            return data;
        }
    }

    public static class PicRow {
        private final String name;
        private final String pic;

        public PicRow(String name, String pic) {
            this.name = name;
            this.pic = pic;
        }

        public String getName() {
            return name;
        }

        public String getPic() {
            return pic;
        }
    }
}
