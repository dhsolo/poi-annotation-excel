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
package io.github.dhsolo.common;

import io.github.dhsolo.image.ImageUtils;
import org.junit.jupiter.api.Test;

import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class ConversionAndUrlTest {

    /**
     * Regression: {@code mapToBean} used to {@code field.set} raw (String) cell values
     * directly, throwing {@code IllegalArgumentException} for any typed field — usable only
     * for all-String beans. Values now convert via the same table as column imports;
     * unconvertible values skip the field instead of failing the whole mapping.
     */
    @Test
    void mapToBeanConvertsStringValuesToFieldTypes() {
        Bean bean = Reflect.mapToBean(Map.of(
                "name", "alpha",
                "age", "42",
                "score", "3.5",
                "active", "true",
                "broken", "not-a-number"), Bean.class);
        assertThat(bean.name).isEqualTo("alpha");
        assertThat(bean.age).isEqualTo(42);
        assertThat(bean.score).isEqualTo(3.5);
        assertThat(bean.active).isTrue();
        assertThat(bean.broken).isNull(); // unconvertible -> skipped, not an exception
    }

    /**
     * Regression: {@code urlEncoder} only encoded the CJK basic block — spaces, full-width
     * punctuation and extension-block characters slipped through unencoded.
     */
    @Test
    void urlEncoderCoversAllIllegalCharacters() throws Exception {
        assertThat(ImageUtils.urlEncoder("http://h/a b.png")).isEqualTo("http://h/a%20b.png");
        assertThat(ImageUtils.urlEncoder("http://h/图 片.png"))
                .isEqualTo("http://h/%E5%9B%BE%20%E7%89%87.png");
        // already-encoded input is not double-encoded; reserved characters survive
        assertThat(ImageUtils.urlEncoder("http://h/a%20b.png?x=1&y=2#f"))
                .isEqualTo("http://h/a%20b.png?x=1&y=2#f");
    }

    /**
     * Regression: the Integer/BigInteger branches cut the numeric string at the first
     * {@code '.'} — for {@code Double.toString} values in scientific notation
     * ({@code "1.2345678E7"}) that silently produced {@code 1} instead of {@code 12345678}.
     */
    @Test
    void convertResolvesScientificNotationForIntegerTargets() {
        assertThat(CommonUtil.convert(12345678.0d, Integer.class)).isEqualTo(12345678);
        assertThat(CommonUtil.convert("1.2E3", java.math.BigInteger.class))
                .isEqualTo(java.math.BigInteger.valueOf(1200));
        // legacy truncation semantics for plain decimals are preserved
        assertThat(CommonUtil.convert("123.99", Integer.class)).isEqualTo(123);
    }

    @Test
    void mapToBeanResolvesDoubleToIntFieldExactly() {
        IntBean bean = Reflect.mapToBean(Map.of("orderId", 12345678.0d), IntBean.class);
        assertThat(bean.orderId).isEqualTo(12345678);
    }

    /**
     * Regression: lenient prefix parsing let {@code "yyyy-MM-dd"} swallow a trailing
     * {@code "HH:mm"} (time silently dropped) and rolled invalid field values over to a
     * different date. Parsing is now strict full-string with in-range fields.
     */
    @Test
    void parseDateIsStrictAndKeepsMinutePrecision() {
        assertThat(CommonUtil.formatDate(CommonUtil.parseDate("2024-05-06 10:30"), "yyyy-MM-dd HH:mm"))
                .isEqualTo("2024-05-06 10:30");
        assertThat(CommonUtil.parseDate("2024-99-99")).isNull();
        assertThat(CommonUtil.formatDate(CommonUtil.parseDate("2024-05-06"), "yyyy-MM-dd"))
                .isEqualTo("2024-05-06");
        assertThat(CommonUtil.formatDate(CommonUtil.parseDate("2024-05-06 10:30:45"), "yyyy-MM-dd HH:mm:ss"))
                .isEqualTo("2024-05-06 10:30:45");
    }

    public static class Bean {
        private String name;
        private Integer age;
        private Double score;
        private Boolean active;
        private Integer broken;

        public Bean() {
        }
    }

    public static class IntBean {
        private int orderId;

        public IntBean() {
        }
    }
}
