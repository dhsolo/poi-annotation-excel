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
 * {@link CSVRow#getPhysicalNumberOfCells()} must report the real cell count (POI contract), not a
 * hard-coded 0, so a CSV sheet handed to generic POI code that iterates by physical count behaves.
 */
class CSVRowCellCountTest {

    @Test
    void physicalCellCountMatchesTokenCount() {
        CSVRow row = new CSVRow(new String[]{"a", "b", "c"}, null, 0);
        assertThat(row.getPhysicalNumberOfCells()).isEqualTo(3);
        assertThat(row.getLastCellNum()).isEqualTo((short) 3); // last index + 1
    }

    @Test
    void emptyRowHasNoCells() {
        CSVRow row = new CSVRow(new String[]{}, null, 0);
        assertThat(row.getPhysicalNumberOfCells()).isZero();
        assertThat(row.getLastCellNum()).isEqualTo((short) -1);
    }
}
