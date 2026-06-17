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
 * Marks a method that provides a dynamic header-to-field-name mapping for import.
 *
 * <p>Apply this annotation on a public no-arg method in the model class that returns a
 * {@code Map<String, String>} where the key is the Excel column header text and the value is the
 * corresponding Java field name. The class-based {@code ExcelUtil.importExcel(in, ..., Class)}
 * overloads read the file's header row and re-order the columns to match by header text rather
 * than by fixed position, so import is robust to column reordering and to header labels that
 * differ from field names; columns whose header is not in the map are skipped. The header row is
 * the row immediately above the first data row ({@code startRow - 1}).
 *
 * <pre>{@code
 * @ExcelHeaderColumnMapping
 * public Map<String, String> headerMapping() {
 *     Map<String, String> map = new LinkedHashMap<>();
 *     map.put("Product Name", "productName");
 *     map.put("Unit Price",   "unitPrice");
 *     return map;
 * }
 * }</pre>
 *
 * @author dhsolo
 * @since 1.0
 */
@Target({ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelHeaderColumnMapping {
}
