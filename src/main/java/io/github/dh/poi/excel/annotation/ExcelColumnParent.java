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
 * Groups multiple {@link ExcelColumn} definitions under a single field, enabling a
 * parent-column / merged-header layout in the generated sheet.
 *
 * <p>Apply this annotation on a field when the sheet requires a two-level header where
 * one parent header cell spans several child column cells. The {@link #columns()} array
 * lists the individual sub-columns that belong to this parent group.
 *
 * <p>Each child {@link ExcelColumn} must declare its value source via
 * {@link ExcelColumn#sourceField()} (or {@link ExcelColumn#sourcePath()}), since the child
 * columns do not map to a Java field directly.
 *
 * <pre>{@code
 * @ExcelColumnParent(value = "销售额", columns = {
 *     @ExcelColumn(columnName = "Q1", index = 2, sourceField = "q1Sales"),
 *     @ExcelColumn(columnName = "Q2", index = 3, sourceField = "q2Sales")
 * })
 * private Object salesGroup;   // anchor field; values come from sourceField on the row object
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumnParent {

    /**
     * The parent header text rendered (merged) above the child columns. Defaults to empty.
     */
    String value() default "";

    /**
     * The child {@link ExcelColumn} definitions that fall under this parent column group.
     * Defaults to an empty array.
     */
    ExcelColumn[] columns() default {};
}
