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
 * Describes the content source for a single cell in a custom row defined by {@link ExcelRow}.
 *
 * <p>This annotation is used as an element inside {@link ExcelRow#cells()} to specify
 * where the cell value comes from. Exactly one of {@link #value()} or {@link #field()}
 * should be set:
 * <ul>
 *   <li>{@link #value()} — a hard-coded static string written into the cell.</li>
 *   <li>{@link #field()} — the name of another field in the same model class whose
 *       runtime value is used.</li>
 * </ul>
 * When neither attribute is set the enclosing {@code @ExcelRow} field's own value is used.
 *
 * <pre>{@code
 * // Static text in the custom row
 * @ExcelRow(cells = @ExcelCell(value = "Remark: draft"))
 * private String ignored;
 *
 * // Value sourced from another field
 * @ExcelRow(cells = @ExcelCell(field = "remark"))
 * private String remarkRow;
 * }</pre>
 *
 * @author dhsolo
 * @since 1.0
 */
@Target({ElementType.FIELD, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCell {

   /**
    * Hard-coded static text to write into the cell.
    * Defaults to an empty string (not used).
    */
   String value() default "";

   /**
    * Name of another field in the same model class whose runtime value is written into
    * the cell. Defaults to an empty string (not used).
    */
   String field() default "";
}
