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
package io.github.dhsolo.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Specifies the date/time format pattern for a date field in an Excel column.
 *
 * <p>Apply this annotation together with {@link ExcelColumn} on a field of type
 * {@link java.util.Date}, {@link java.time.LocalDate}, or {@link java.time.LocalDateTime}.
 * During export the field value is formatted with the given pattern; during import the
 * cell string is parsed using the same pattern.
 *
 * <pre>{@code
 * @ExcelColumn(columnName = "Created At", index = 5)
 * @ExcelDateFormat(pattern = "yyyy/MM/dd")
 * private Date createdAt;
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelDateFormat {

    /**
     * Date/time format pattern compatible with {@link java.text.SimpleDateFormat}.
     * Defaults to {@code "yyyy-MM-dd HH:mm:ss"}.
     */
    String pattern() default "yyyy-MM-dd HH:mm:ss";
}
