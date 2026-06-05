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
 * Marks a field or method that contributes one custom (non-data) row to the generated sheet.
 *
 * <p>Apply this annotation on one or more fields in a model class annotated with
 * {@link ExcelInfo} to inject fully merged rows (spanning all columns) at a fixed
 * position above the data rows — for example, to add sub-headings, remark lines, or
 * summary rows. Use the {@link #order()} attribute to control the relative order when
 * multiple such fields exist.
 *
 * <p>Cell content is resolved as follows (see {@link ExcelCell} for details):
 * <ul>
 *   <li>No {@link #cells()} specified: the field's own string value is written.</li>
 *   <li>{@code cells = @ExcelCell(value = "...")}: a static string is written.</li>
 *   <li>{@code cells = @ExcelCell(field = "...")}: the value of the named sibling field is written.</li>
 * </ul>
 * Rows whose resolved content is {@code null} or blank are silently skipped.
 *
 * <pre>{@code
 * @ExcelRow(order = 1)
 * private String remark = "Internal use only";
 *
 * @ExcelRow(order = 2, cells = @ExcelCell(field = "projectName"))
 * private String projectRow;
 * }</pre>
 *
 * @author dhsolo
 * @since 1.0
 */
@Target({ElementType.FIELD, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelRow {

    /**
     * Cell content configuration.
     * <ul>
     *   <li>Not provided (default): uses the annotated field's own value as the row content, merging the entire row.</li>
     *   <li>{@code @ExcelCell(value = "static text")}: uses static text.</li>
     *   <li>{@code @ExcelCell(field = "otherField")}: uses the value of another field in the same class.</li>
     * </ul>
     */
    ExcelCell[] cells() default {};

    /**
     * Display order when multiple rows are present; smaller values appear first. Defaults to 0.
     *
     * <p><b>Important:</b> When a class has multiple {@code @ExcelRow} fields,
     * each field must specify a distinct {@code order} value to guarantee deterministic ordering.
     * If multiple fields share the same {@code order}, their relative order depends on the JVM's
     * {@code getDeclaredFields()} implementation and is not guaranteed by the specification.
     *
     * <pre>
     * // Correct: unique order values, deterministic order
     * {@literal @}ExcelRow(order = 1)
     * private String schemeName;
     *
     * {@literal @}ExcelRow(order = 2)
     * private String areaName;
     *
     * // Incorrect: same order, non-deterministic order
     * {@literal @}ExcelRow
     * private String schemeName;
     * {@literal @}ExcelRow
     * private String areaName;
     * </pre>
     */
    int order() default 0;
}
