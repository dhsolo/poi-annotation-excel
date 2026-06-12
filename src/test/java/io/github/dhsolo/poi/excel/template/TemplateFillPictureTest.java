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

import com.sun.net.httpserver.HttpServer;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.net.InetSocketAddress;
import java.net.ServerSocket;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Picture placeholders (issue #5): {@code ${@image:key}} anchors an image registered via
 * {@code fillPicture} over the placeholder cell, {@code ${list.@image:key}} takes the image
 * from each row map. Placeholder text is cleared, the sniffed format is preserved, and the
 * anchors cooperate with the list-expansion row shift.
 */
class TemplateFillPictureTest {

    @Test
    void scalarPlaceholderAnchorsPictureAndClearsText() throws Exception {
        byte[] template = template(sheet -> sheet.createRow(0).createCell(1).setCellValue("${@image:logo}"));

        byte[] filled = ExcelTemplateFiller.of(template)
                .fillPicture("logo", png())
                .toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEmpty();
            assertThat(wb.getAllPictures()).hasSize(1);
            assertThat(wb.getAllPictures().get(0).suggestFileExtension()).isEqualTo("png");
            List<XSSFPicture> pictures = pictures(sheet);
            assertThat(pictures).hasSize(1);
            ClientAnchor anchor = pictures.get(0).getClientAnchor();
            assertThat(anchor.getRow1()).isZero();
            assertThat(anchor.getCol1()).isEqualTo((short) 1);
            assertThat(anchor.getRow2()).isEqualTo(1);
            assertThat(anchor.getCol2()).isEqualTo((short) 2);
        }
    }

    @Test
    void jpegBytesKeepJpegFormat() throws Exception {
        byte[] template = template(sheet -> sheet.createRow(0).createCell(0).setCellValue("${@image:photo}"));

        byte[] filled = ExcelTemplateFiller.of(template)
                .fillPicture("photo", jpeg())
                .toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            assertThat(wb.getAllPictures()).hasSize(1);
            assertThat(wb.getAllPictures().get(0).suggestFileExtension()).isEqualTo("jpeg");
        }
    }

    @Test
    void listRowsGetOnePicturePerRowFromRowData() throws Exception {
        byte[] template = template(sheet -> {
            var row = sheet.createRow(2);
            row.createCell(0).setCellValue("${list.name}");
            row.createCell(1).setCellValue("${list.@image:photo}");
        });

        List<Map<String, Object>> rows = new ArrayList<>();
        for (String name : List.of("a", "b", "c")) {
            Map<String, Object> row = new LinkedHashMap<>();
            row.put("name", name);
            row.put("photo", png());
            rows.add(row);
        }

        byte[] filled = ExcelTemplateFiller.of(template).fillList("list", rows).toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("a");
            assertThat(sheet.getRow(4).getCell(0).getStringCellValue()).isEqualTo("c");
            assertThat(sheet.getRow(3).getCell(1).getStringCellValue()).isEmpty();
            List<Integer> anchorRows = pictures(sheet).stream().map(p -> p.getClientAnchor().getRow1()).toList();
            assertThat(anchorRows).containsExactlyInAnyOrder(2, 3, 4);
        }
    }

    @Test
    void rowDataAcceptsFileAndStreamValuesAndSkipsMissingOnes() throws Exception {
        byte[] template = template(sheet -> {
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue("${list.name}");
            row.createCell(1).setCellValue("${list.@image:photo}");
        });

        File pngFile = File.createTempFile("fill-picture", ".png");
        pngFile.deleteOnExit();
        Files.write(pngFile.toPath(), png());

        List<Map<String, Object>> rows = new ArrayList<>();
        Map<String, Object> fromFile = new LinkedHashMap<>();
        fromFile.put("name", "file");
        fromFile.put("photo", pngFile);
        rows.add(fromFile);
        Map<String, Object> fromStream = new LinkedHashMap<>();
        fromStream.put("name", "stream");
        fromStream.put("photo", new ByteArrayInputStream(png()));
        rows.add(fromStream);
        Map<String, Object> missing = new LinkedHashMap<>();
        missing.put("name", "missing"); // no photo entry: placeholder cleared, no picture
        rows.add(missing);

        byte[] filled = ExcelTemplateFiller.of(template).fillList("list", rows).toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEmpty();
            List<Integer> anchorRows = pictures(sheet).stream().map(p -> p.getClientAnchor().getRow1()).toList();
            assertThat(anchorRows).containsExactlyInAnyOrder(0, 1);
        }
    }

    @Test
    void scalarPlaceholderBelowExpandedListShiftsWithItsRow() throws Exception {
        byte[] template = template(sheet -> {
            sheet.createRow(1).createCell(0).setCellValue("${list.name}");
            sheet.createRow(4).createCell(0).setCellValue("${@image:logo}");
        });

        byte[] filled = ExcelTemplateFiller.of(template)
                .fillList("list", List.of(Map.of("name", "a"), Map.of("name", "b"), Map.of("name", "c")))
                .fillPicture("logo", png())
                .toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            Sheet sheet = wb.getSheetAt(0);
            // list grew by 2 rows, so the placeholder cell moved 4 -> 6 and the picture followed
            List<Integer> anchorRows = pictures(sheet).stream().map(p -> p.getClientAnchor().getRow1()).toList();
            assertThat(anchorRows).containsExactly(6);
            assertThat(sheet.getRow(6).getCell(0).getStringCellValue()).isEmpty();
        }
    }

    @Test
    void missingKeyAndUnrecognizableBytesClearPlaceholderWithoutPicture() throws Exception {
        byte[] template = template(sheet -> {
            sheet.createRow(0).createCell(0).setCellValue("${@image:absent}");
            sheet.createRow(1).createCell(0).setCellValue("${@image:garbage}");
        });

        byte[] filled = ExcelTemplateFiller.of(template)
                .fillPicture("garbage", "not-an-image".getBytes(StandardCharsets.UTF_8))
                .toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEmpty();
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEmpty();
            assertThat(wb.getAllPictures()).isEmpty();
            assertThat(sheet.getDrawingPatriarch()).isNull();
        }
    }

    @Test
    void imagePlaceholderCoexistsWithTextInTheSameCell() throws Exception {
        byte[] template = template(sheet -> sheet.createRow(0).createCell(0)
                .setCellValue("${@image:logo}${title}"));

        byte[] filled = ExcelTemplateFiller.of(template)
                .fillPicture("logo", png())
                .fill("title", "Report")
                .toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Report");
            assertThat(pictures(sheet)).hasSize(1);
        }
    }

    @Test
    void urlValueIsDownloadedThroughTheGuardedFetch() throws Exception {
        byte[] template = template(sheet -> sheet.createRow(0).createCell(0).setCellValue("${@image:logo}"));
        AtomicInteger hits = new AtomicInteger();
        HttpServer server = pngServer(hits);
        try {
            String url = "http://127.0.0.1:" + server.getAddress().getPort() + "/logo.png";

            byte[] filled = ExcelTemplateFiller.of(template).fillPicture("logo", url).toBytes();

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
                assertThat(wb.getAllPictures()).hasSize(1);
                assertThat(wb.getAllPictures().get(0).suggestFileExtension()).isEqualTo("png");
                assertThat(pictures(wb.getSheetAt(0))).hasSize(1);
            }
            assertThat(hits).hasValue(1);
        } finally {
            server.stop(0);
        }
    }

    @Test
    void sameUrlAcrossRowsDownloadsOnceAndSharesOneMediaPart() throws Exception {
        byte[] template = template(sheet -> {
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue("${list.name}");
            row.createCell(1).setCellValue("${list.@image:photo}");
        });
        AtomicInteger hits = new AtomicInteger();
        HttpServer server = pngServer(hits);
        try {
            String url = "http://127.0.0.1:" + server.getAddress().getPort() + "/logo.png";
            List<Map<String, Object>> rows = List.of(
                    Map.of("name", "a", "photo", url),
                    Map.of("name", "b", "photo", url),
                    Map.of("name", "c", "photo", url));

            byte[] filled = ExcelTemplateFiller.of(template).fillList("list", rows).toBytes();

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
                Sheet sheet = wb.getSheetAt(0);
                assertThat(wb.getAllPictures()).hasSize(1); // one media part, anchored three times
                List<Integer> anchorRows = pictures(sheet).stream().map(p -> p.getClientAnchor().getRow1()).toList();
                assertThat(anchorRows).containsExactlyInAnyOrder(0, 1, 2);
            }
            assertThat(hits).hasValue(1);
        } finally {
            server.stop(0);
        }
    }

    @Test
    void unreachableUrlClearsPlaceholdersWithoutFailingTheFill() throws Exception {
        byte[] template = template(sheet -> {
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue("${list.name}");
            row.createCell(1).setCellValue("${list.@image:photo}");
        });
        int deadPort;
        try (ServerSocket socket = new ServerSocket(0)) {
            deadPort = socket.getLocalPort();
        }
        String deadUrl = "http://127.0.0.1:" + deadPort + "/gone.png";
        List<Map<String, Object>> rows = List.of(
                Map.of("name", "a", "photo", deadUrl),
                Map.of("name", "b", "photo", deadUrl));

        byte[] filled = ExcelTemplateFiller.of(template).fillList("list", rows).toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(wb.getAllPictures()).isEmpty();
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEmpty();
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("b");
        }
    }

    @Test
    void stringValueWithoutHttpPrefixReadsALocalFile() throws Exception {
        byte[] template = template(sheet -> sheet.createRow(0).createCell(0).setCellValue("${@image:logo}"));
        File pngFile = File.createTempFile("fill-picture-url", ".png");
        pngFile.deleteOnExit();
        Files.write(pngFile.toPath(), png());

        byte[] filled = ExcelTemplateFiller.of(template)
                .fillPicture("logo", pngFile.getAbsolutePath())
                .toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(filled))) {
            assertThat(wb.getAllPictures()).hasSize(1);
            assertThat(wb.getAllPictures().get(0).suggestFileExtension()).isEqualTo("png");
        }
    }

    private static HttpServer pngServer(AtomicInteger hits) throws IOException {
        byte[] body = image("png");
        HttpServer server = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        server.createContext("/logo.png", exchange -> {
            hits.incrementAndGet();
            exchange.getResponseHeaders().add("Content-Type", "image/png");
            exchange.sendResponseHeaders(200, body.length);
            try (OutputStream os = exchange.getResponseBody()) {
                os.write(body);
            }
        });
        server.start();
        return server;
    }

    private static List<XSSFPicture> pictures(Sheet sheet) {
        List<XSSFPicture> result = new ArrayList<>();
        XSSFDrawing drawing = (XSSFDrawing) sheet.getDrawingPatriarch();
        if (drawing == null) return result;
        for (var shape : drawing.getShapes()) {
            if (shape instanceof XSSFPicture pic) result.add(pic);
        }
        return result;
    }

    private interface SheetLayout {
        void apply(XSSFSheet sheet);
    }

    private static byte[] template(SheetLayout layout) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            layout.apply(wb.createSheet("t"));
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] png() throws IOException {
        return image("png");
    }

    private static byte[] jpeg() throws IOException {
        return image("jpg");
    }

    private static byte[] image(String format) throws IOException {
        BufferedImage img = new BufferedImage(6, 6, BufferedImage.TYPE_INT_RGB);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ImageIO.write(img, format, out);
        return out.toByteArray();
    }
}
