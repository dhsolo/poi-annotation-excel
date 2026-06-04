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

import io.github.dh.poi.excel.core.ContextType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Marks a field or method that injects extra contextual rows at a specific position
 * within the generated sheet.
 *
 * <p>The {@link #type()} attribute specifies where the contextual content is inserted
 * relative to the structural sections of the sheet:
 * <ul>
 *   <li>{@code BEFORE_TITLE} — before the title row.</li>
 *   <li>{@code BETWEEN_TITLE_HEADER} — between the title row and the header row.</li>
 *   <li>{@code BETWEEN_HEADER_DATA} — between the header row and the first data row.</li>
 *   <li>{@code AFTER_DATA} — after the last data row.</li>
 * </ul>
 *
 * <pre>{@code
 * @ExcelContext(type = ContextType.BETWEEN_TITLE_HEADER)
 * private String subTitle = "Generated on 2024-01-01";
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
@Target({ElementType.FIELD, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelContext {

     /**
      * The position within the sheet at which the contextual content is inserted.
      * Must be one of the {@link ContextType} enum values; this attribute is required.
      */
     ContextType type();
}
