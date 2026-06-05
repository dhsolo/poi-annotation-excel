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
 * Attaches a dropdown validation list to an Excel column.
 *
 * <p>Can be placed on a field (together with {@link ExcelColumn}) to declare a static
 * list of allowed values, or on a method to supply the list dynamically at runtime.
 *
 * <p><b>Field usage</b> — inline static list:
 * <pre>{@code
 * @ExcelColumn(columnName = "Status", index = 2)
 * @ExcelListBox(listTextBox = {"Pending", "Approved", "Rejected"})
 * private String status;
 * }</pre>
 *
 * <p><b>Method usage</b> — dynamic list (method must return {@code List<String>},
 * {@code Set<String>}, {@code String[]}, or {@code CascadeValidateModel}):
 * <pre>{@code
 * @ExcelListBox(columnName = "status")
 * public List<String> statusOptions() {
 *     return Arrays.asList("Pending", "Approved", "Rejected");
 * }
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
@Target({ElementType.FIELD, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelListBox {

    /**
     * Static array of allowed dropdown values.
     * Only meaningful when placed on a field; ignored when placed on a method.
     * Defaults to an empty array.
     */
    String[] listTextBox() default {};

    /**
     * Name of the target column field this list applies to.
     * Required when the annotation is placed on a method; must match the field name
     * of the corresponding {@link ExcelColumn}-annotated field.
     * Defaults to an empty string.
     */
    String columnName() default "";

    /**
     * Whether to throw a validation exception when the imported cell value is not
     * found in the dropdown list. Set to {@code false} to silently accept unlisted values.
     * Defaults to {@code true}.
     */
    boolean isNeedAddTranslationException() default true;
}
