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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: the synchronous fallback in {@link DefaultPictureHandler#setPicture} (taken when
 * the disk-staging directory is unavailable) used to open the URL directly, bypassing the
 * protocol whitelist, {@link ImageDownloadPolicy} SSRF guard, and the download size cap that
 * the async path enforces. It now goes through the same guarded stream.
 */
class FallbackPathPolicyTest {

    private final ThreadPoolExecutor executor =
            new ThreadPoolExecutor(1, 1, 0, TimeUnit.SECONDS, new LinkedBlockingQueue<>());

    @AfterEach
    void reset() {
        ImageDownloadPolicy.setBlockPrivateNetworks(false);
        executor.shutdownNow();
    }

    @Test
    void fallbackPathHonoursPrivateNetworkBlock() throws Exception {
        ImageDownloadPolicy.setBlockPrivateNetworks(true);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("s");
            Cell cell = sheet.createRow(0).createCell(0);
            DefaultPictureHandler handler = new DefaultPictureHandler(
                    wb, "xlsx", sheet, sheet.createDrawingPatriarch(), ",", executor, null);
            // pictureDirCreate stays false -> synchronous fallback path
            int extra = handler.setPicture(0, 0, 0, 0, "http://127.0.0.1/x.png", cell);

            assertThat(extra).isEqualTo(0);
            // Blocked image is skipped (logged) and no picture is embedded; the cell is left blank
            // rather than stamped with error text.
            assertThat(cell.getStringCellValue()).isEmpty();
            assertThat(wb.getAllPictures()).isEmpty();
        }
    }
}
