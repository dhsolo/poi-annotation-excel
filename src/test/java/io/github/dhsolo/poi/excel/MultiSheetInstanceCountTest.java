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
package io.github.dhsolo.poi.excel;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression: a multi-sheet export stitches the 2nd..Nth creators into the first and only closes
 * the first, so without un-counting the stitched children {@code INSTANCE_COUNT} would creep up by
 * (sheets - 1) every call and the shared download pool's graceful shutdown would never fire again.
 */
class MultiSheetInstanceCountTest {

    @Test
    void multiSheetExportBalancesInstanceCount() throws Exception {
        int before = instanceCount();

        ExcelUtil.toBytes(
                ExcelCreatorBuilder.create("s1").columns("ID:id").data(List.of(Map.of("id", "1"))),
                ExcelCreatorBuilder.create("s2").columns("ID:id").data(List.of(Map.of("id", "2"))),
                ExcelCreatorBuilder.create("s3").columns("ID:id").data(List.of(Map.of("id", "3"))));

        // Every creator the call constructed must have been accounted for (the two stitched child
        // sheets used to leak, leaving the count at before + 2).
        assertThat(instanceCount()).isEqualTo(before);
    }

    private static int instanceCount() throws Exception {
        Field f = ExcelCreator.class.getDeclaredField("INSTANCE_COUNT");
        f.setAccessible(true);
        return ((AtomicInteger) f.get(null)).get();
    }
}
