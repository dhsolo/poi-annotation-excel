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
 * Marks the field or method that provides the title text rendered above the header row.
 *
 * <p>Apply this annotation to a {@code String} field (or a no-arg method returning
 * {@code String}) on a class annotated with {@link ExcelInfo}. The framework reads
 * the value at runtime and writes it as a merged title row spanning all data columns.
 * If the field value is {@code null} or empty, no title row is emitted.
 *
 * <pre>{@code
 * @ExcelInfo(sheetName = "Sales")
 * public class SalesExcelModel {
 *     @ExcelTitle
 *     private String title = "Monthly Sales Report";
 * }
 * }</pre>
 *
 * @author dhsolo
 * @since 1.0
 */
@Target({ElementType.FIELD, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelTitle {
}
