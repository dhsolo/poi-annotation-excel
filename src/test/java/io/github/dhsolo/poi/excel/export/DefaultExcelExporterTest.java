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
package io.github.dhsolo.poi.excel.export;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/** Covers the exporter idempotency (zip mode) and local-file extension handling fixes. */
class DefaultExcelExporterTest {

    private static byte[] workbookBytes() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            wb.createSheet("s");
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static Workbook workbookWithSheet() {
        XSSFWorkbook wb = new XSSFWorkbook();
        wb.createSheet("s");
        return wb;
    }

    @Test
    void zipModeReadsAreIdempotentAndTempFileSurvivesUntilClose() throws Exception {
        File temp = File.createTempFile("exporter", ".xlsx");
        Files.write(temp.toPath(), workbookBytes());

        DefaultExcelExporter exporter = new DefaultExcelExporter(new XSSFWorkbook(), "xlsx");
        exporter.setZip(true);
        exporter.setTempWorkFile(temp);

        byte[] first;
        try (InputStream in = exporter.getInputStream(true)) {
            first = in.readAllBytes();
        }
        // The second read used to NPE because the first call deleted the temp file in finally.
        byte[] second;
        try (InputStream in = exporter.getInputStream(true)) {
            second = in.readAllBytes();
        }
        assertThat(second).isEqualTo(first);
        assertThat(temp).exists();

        exporter.close();
        assertThat(temp).doesNotExist();
    }

    @Test
    void exportLocalDoesNotDoubleAppendExtensionCaseInsensitively() throws Exception {
        Path dir = Files.createTempDirectory("exp");
        DefaultExcelExporter exporter = new DefaultExcelExporter(workbookWithSheet(), "xlsx");

        String upper = dir.resolve("report.XLSX").toString();
        exporter.exportLocal(upper);

        assertThat(new File(upper)).exists();                  // report.XLSX — not doubled
        assertThat(new File(upper + ".xlsx")).doesNotExist();
        exporter.close();
    }

    @Test
    void exportLocalAppendsExtensionWhenMissing() throws Exception {
        Path dir = Files.createTempDirectory("exp");
        DefaultExcelExporter exporter = new DefaultExcelExporter(workbookWithSheet(), "xlsx");

        exporter.exportLocal(dir.resolve("noext").toString());

        assertThat(new File(dir.resolve("noext.xlsx").toString())).exists();
        exporter.close();
    }
}
