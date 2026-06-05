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
 * Associates an Excel formula with a column field.
 *
 * <p>Apply this annotation together with {@link ExcelColumn} on a field to have the
 * framework write a formula string into each data cell of that column rather than a
 * plain value. The formula follows standard Excel syntax and may use relative row
 * references which the framework resolves at write time.
 *
 * <pre>{@code
 * @ExcelColumn(columnName = "Total", index = 4)
 * @ExcelFormula(formula = "B{row}*C{row}")
 * private BigDecimal total;
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelFormula {

   /**
    * The Excel formula string to set on each data cell. Use {@code {row}} as a
    * placeholder for the current one-based row number if row-relative references are
    * needed. This attribute is required.
    */
   String formula();
}
