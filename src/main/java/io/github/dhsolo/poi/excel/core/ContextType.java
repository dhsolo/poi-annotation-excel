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
package io.github.dhsolo.poi.excel.core;

/**
 * Enumeration of the positional context zones within a generated Excel sheet.
 *
 * <p>The export pipeline divides a sheet into four logical zones relative to the
 * structural landmarks of a typical report layout:</p>
 * <ol>
 *   <li>{@link #BEFORE_TITLE} - rows inserted before the title row</li>
 *   <li>{@link #BETWEEN_TITLE_HEADER} - rows between the title and the column-header row</li>
 *   <li>{@link #BETWEEN_HEADER_DATA} - rows between the column-header row and the first data row</li>
 *   <li>{@link #AFTER_DATA} - rows appended after the last data row</li>
 * </ol>
 * <p>These values are used by context-aware hooks or renderers that need to inject content
 * at a specific structural position in the sheet.</p>
 *
 * @author dhsolo
 * @since 1.0
 */
public enum ContextType {

    /** Zone before the title row. */
    BEFORE_TITLE,

    /** Zone between the title row and the column-header row. */
    BETWEEN_TITLE_HEADER,

    /** Zone between the column-header row and the first data row. */
    BETWEEN_HEADER_DATA,

    /** Zone after the last data row. */
    AFTER_DATA
}
