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
package io.github.dh.poi.excel;

import io.github.dh.poi.excel.annotation.ExcelColumn;
import io.github.dh.poi.excel.annotation.ExcelColumnParent;
import io.github.dh.poi.excel.annotation.ExcelData;
import io.github.dh.poi.excel.annotation.ExcelInfo;
import io.github.dh.poi.excel.exception.ExcelAnnotationException;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThatThrownBy;

/** A @ExcelColumnParent group whose columns are interrupted by another column must fail fast. */
class GroupedHeaderValidationTest {

    @Test
    void nonContiguousGroupFailsFast() {
        assertThatThrownBy(() -> ExcelUtil.toBytes(new BadModel().setData(List.of(new R("a", 1, 2, "x")))))
                .isInstanceOf(ExcelAnnotationException.class)
                .hasMessageContaining("non-contiguous");
    }

    @ExcelInfo(sheetName = "bad", isBigData = false)
    public static class BadModel {
        @ExcelData
        private List<R> data;

        // Group "G" occupies index 1 and 3, interrupted by "中间列" at index 2.
        @ExcelColumnParent(value = "G", columns = {
                @ExcelColumn(columnName = "A", index = 1, sourceField = "a"),
                @ExcelColumn(columnName = "B", index = 3, sourceField = "b")
        })
        private Object group;

        @ExcelColumn(columnName = "中间列", index = 2)
        private String middle;

        public BadModel setData(List<R> data) { this.data = data; return this; }
        public List<R> getData() { return data; }
    }

    public static class R {
        private final Integer a;
        private final Integer b;
        private final String middle;

        public R(String middle, Integer a, Integer b, String ignored) {
            this.middle = middle;
            this.a = a;
            this.b = b;
        }

        public Integer getA() { return a; }
        public Integer getB() { return b; }
        public String getMiddle() { return middle; }
    }
}
