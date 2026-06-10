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
package io.github.dhsolo.poi.excel.template;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: list-row expansion shifts cells via {@code shiftRows}, but POI does not move
 * drawing anchors — a picture anchored below the list region used to stay put and overlap the
 * inserted rows. Anchors below the template row are now shifted by the same amount; pictures
 * above the list region stay where they are.
 */
class TemplatePictureAnchorTest {

    @Test
    void picturesBelowExpandedListShiftWithTheRows() throws Exception {
        byte[] template = templateWithPictures();

        byte[] filled = ExcelTemplateFiller.of(template)
                .fillList("list", List.of(Map.of("name", "a"), Map.of("name", "b"), Map.of("name", "c")))
                .toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(wb.getAllPictures()).hasSize(1); // one media part, anchored twice
            // the marker text that sat next to the lower picture moved from row 5 to row 7
            assertThat(sheet.getRow(7).getCell(0).getStringCellValue()).isEqualTo("MARKER");

            List<Integer> anchorRows = anchorRows(sheet);
            // header picture (row 0) untouched; lower picture follows its row: 5 -> 7
            assertThat(anchorRows).containsExactlyInAnyOrder(0, 7);
        }
    }

    private static List<Integer> anchorRows(Sheet sheet) {
        List<Integer> rows = new ArrayList<>();
        XSSFDrawing drawing = (XSSFDrawing) sheet.getDrawingPatriarch();
        for (var shape : drawing.getShapes()) {
            if (shape instanceof XSSFPicture pic) {
                ClientAnchor a = pic.getPreferredSize();
                rows.add(a.getRow1());
            }
        }
        return rows;
    }

    /** Row 2 = list placeholder; pictures anchored at row 0 (above) and row 5 (below) + MARKER text. */
    private static byte[] templateWithPictures() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            XSSFSheet sheet = wb.createSheet("t");
            sheet.createRow(2).createCell(0).setCellValue("${list.name}");
            sheet.createRow(5).createCell(0).setCellValue("MARKER");

            BufferedImage img = new BufferedImage(6, 6, BufferedImage.TYPE_INT_RGB);
            ByteArrayOutputStream png = new ByteArrayOutputStream();
            ImageIO.write(img, "png", png);
            int picIdx = wb.addPicture(png.toByteArray(), Workbook.PICTURE_TYPE_PNG);
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            drawing.createPicture(new XSSFClientAnchor(0, 0, 0, 0, 1, 0, 2, 1), picIdx); // header pic, row 0
            drawing.createPicture(new XSSFClientAnchor(0, 0, 0, 0, 1, 5, 2, 6), picIdx); // below list, row 5
            wb.write(out);
            return out.toByteArray();
        }
    }
}
