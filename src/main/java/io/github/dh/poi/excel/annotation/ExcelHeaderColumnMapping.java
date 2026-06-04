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
 * Marks a method that provides a dynamic header-to-field-name mapping for import.
 *
 * <p>Apply this annotation on a no-arg method in the model class that returns a
 * {@code Map<String, String>} where the key is the Excel column header text and the
 * value is the corresponding Java field name of the row data class. This mapping is
 * used during import to match columns by header label rather than by fixed index,
 * making the import robust against column reordering in the source file.
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
 * @author dh
 * @since 1.0
 */
@Target({ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelHeaderColumnMapping {
}
