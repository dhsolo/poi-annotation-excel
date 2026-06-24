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

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatCode;

/**
 * Covers the {@link ExcelTemplateFiller} ownership contract: a filler created from a stream/bytes
 * owns and closes its workbook (usable in try-with-resources), while a filler wrapping a
 * caller-supplied {@link Workbook} leaves that workbook open on {@link ExcelTemplateFiller#close()}.
 */
class TemplateFillerCloseTest {

    private static byte[] templateBytes(String... cells) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("Sheet1");
            Row row = sheet.createRow(0);
            for (int i = 0; i < cells.length; i++) {
                row.createCell(i).setCellValue(cells[i]);
            }
            wb.write(out);
            return out.toByteArray();
        }
    }

    @Test
    void ownedFillerUsableInTryWithResources() throws Exception {
        byte[] template = templateBytes("${name}");
        byte[] result;
        try (ExcelTemplateFiller filler = ExcelTemplateFiller.of(template)) {
            result = filler.fill("name", "Alice").toBytes();
        }
        try (Workbook out = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            assertThat(out.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo("Alice");
        }
    }

    @Test
    void closeLeavesCallerSuppliedWorkbookOpen() throws Exception {
        Workbook caller = new XSSFWorkbook(new ByteArrayInputStream(templateBytes("${name}")));

        ExcelTemplateFiller.of(caller).fill("name", "Bob").toBytes();
        ExcelTemplateFiller.of(caller).close();

        // If close() had closed the caller's workbook, writing it out would throw.
        assertThatCode(() -> caller.write(new ByteArrayOutputStream())).doesNotThrowAnyException();
        caller.close();
    }
}
