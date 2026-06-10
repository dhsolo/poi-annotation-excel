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

import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Instance accounting and streaming-workbook cleanup.
 *
 * <p>Regressions guarded here: a wrapper creator built via {@code ExcelCreator(Workbook)} was
 * never counted but its {@code close()} still decremented the shared instance count, driving it
 * negative and shutting the shared image-download pool down underneath other active exports;
 * a double {@code close()} decremented twice with the same effect.
 */
class ResourceCleanupTest {

    private static int instanceCount() throws Exception {
        Field f = ExcelCreator.class.getDeclaredField("INSTANCE_COUNT");
        f.setAccessible(true);
        return ((AtomicInteger) f.get(null)).get();
    }

    @Test
    void wrapperCreatorCloseDoesNotTouchInstanceCount() throws Exception {
        try (ExcelCreator owner = new ExcelCreator("s1")) {
            int counted = instanceCount();
            ExcelCreator wrapper = new ExcelCreator(owner.getWorkBook());
            wrapper.close();
            assertThat(instanceCount()).isEqualTo(counted);
        }
    }

    @Test
    void doubleCloseDecrementsOnce() throws Exception {
        ExcelCreator creator = new ExcelCreator("s1");
        int counted = instanceCount();
        creator.close();
        assertThat(instanceCount()).isEqualTo(counted - 1);
        creator.close();
        assertThat(instanceCount()).isEqualTo(counted - 1);
    }

    /**
     * Regression: the multi-sheet assembly paths in ExcelUtil used to close only the first
     * creator, leaving the children counted in the shared pool accounting forever.
     */
    @Test
    void multiSheetExportClosesEveryCreator() throws Exception {
        int before = instanceCount();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelUtil.export(out, "multi.xlsx",
                sheetCreator("s1"), sheetCreator("s2"), sheetCreator("s3"));
        assertThat(out.size()).isGreaterThan(0);
        assertThat(instanceCount()).isEqualTo(before);
    }

    private static ExcelCreator sheetCreator(String name) {
        ExcelCreator creator = new ExcelCreator("xlsx", name, false);
        creator.setHeader(new String[]{"名称"});
        Map<Integer, ExcelModel> mapping = new java.util.HashMap<>();
        mapping.put(0, ExcelCreator.generate("name"));
        creator.setColumnMappingInfo(mapping);
        creator.setObject(List.of(new Bean("a")));
        return creator;
    }

    @Test
    void bigDataExportAndCloseDisposesWithoutError() throws Exception {
        ExcelCreator creator = new ExcelCreator("xlsx", "big", true);
        creator.setHeader(new String[]{"名称"});
        Map<Integer, ExcelModel> mapping = new java.util.HashMap<>();
        mapping.put(0, ExcelCreator.generate("name"));
        creator.setColumnMappingInfo(mapping);
        creator.setObject(List.of(new Bean("a"), new Bean("b")));
        creator.createExcel();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        creator.export(out, "big.xlsx");
        assertThat(out.size()).isGreaterThan(0);
        creator.close();
        creator.close(); // idempotent, including the streaming dispose
    }

    public static class Bean {
        private final String name;

        public Bean(String name) {
            this.name = name;
        }

        public String getName() { return name; }
    }
}
