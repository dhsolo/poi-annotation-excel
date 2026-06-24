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
package io.github.dhsolo.poi.excel.core;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: picture pre-registration on POI 5.5 deduped by wrapper identity
 * ({@code List.contains} with no {@code equals} on {@code XSSFPictureData}), so the first
 * registered part — already wrapped by the lazy {@code getAllPictures()} scan — was added a
 * second time, shifting every later by-index anchor to the wrong image. Dedup now compares
 * package part names.
 */
class PicturePreRegistrationTest {

    @Test
    void preRegisteredPicturesAreNotDuplicatedAndKeepTheirOrder() throws Exception {
        try (BusinessXSSFWorkbook wb = new BusinessXSSFWorkbook()) {
            wb.addPicture(1);
            wb.addPicture(2);

            // without part-name dedup this was [image1, image1, image2]: the lazy scan
            // had already wrapped image1, identity contains() missed it and re-added it
            assertThat(wb.getAllPictures()).hasSize(2);
            assertThat(wb.getAllPictures().get(0).getPackagePart().getPartName().getName())
                    .contains("image1");
            assertThat(wb.getAllPictures().get(1).getPackagePart().getPartName().getName())
                    .contains("image2");
        }
    }
}
