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
package io.github.dh.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Marks a method that supplies a custom value-translation function for a specific column.
 *
 * <p>Apply this annotation on a method in the model class to override the default cell
 * value resolution for the column identified by {@link #columnName()}. The annotated
 * method must return a {@code Function<ExcelRowData, Object>}; the framework calls that
 * function for each data row and writes the returned object into the column cell.
 * This takes higher priority than both the inline {@code translate()} mapping and the
 * {@code sourcePath} / {@code sourceField} attributes on {@link ExcelColumn}.
 *
 * <pre>{@code
 * @ExcelTranslateMethod(columnName = "status")
 * public Function<ExcelRowData, Object> translateStatus() {
 *     return row -> row.getStatus() == 1 ? "Active" : "Inactive";
 * }
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
@Target({ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelTranslateMethod {

    /**
     * Name of the column field (i.e. the Java field name of the corresponding
     * {@link ExcelColumn}-annotated field) whose cell value this method translates.
     * This attribute is required.
     */
    String columnName();
}
