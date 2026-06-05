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
package io.github.dhsolo.poi.excel.render;

import io.github.dhsolo.poi.excel.model.ExcelRowData;

/**
 * Extension point for transforming or translating a cell value before it is written
 * to the Excel output during export.
 *
 * <p>Implement this interface to apply cross-cutting value transformations that are
 * independent of any single column definition — for example, replacing internal code
 * values with display labels, formatting numbers according to locale rules, or computing
 * a derived cell value from multiple fields of the source object.
 *
 * <p>The framework calls {@link #needHandle()} first; if it returns {@code false} the
 * handler is skipped entirely and the default cell-value logic applies.  Returning
 * {@code true} signals that {@link #handler(ExcelRowData)} should be called to produce
 * the replacement value.
 *
 * <p>Both methods have default implementations so that implementors only need to
 * override the methods relevant to their use-case.
 *
 * @author dh
 * @since 1.0
 */
public interface ExcelTranslateHandler {

    /**
     * Returns whether this handler should intercept the current cell.
     *
     * <p>Implementations can inspect static context (e.g. a per-column flag set at
     * construction time) or delegate to {@link ExcelRowData} to make a row-specific
     * decision.  The default implementation always returns {@code false}, meaning the
     * handler is opt-in and inactive unless overridden.
     *
     * @return {@code true} if {@link #handler(ExcelRowData)} should be invoked to
     *         produce the cell value; {@code false} to skip this handler
     */
    default boolean needHandle() {
        return false;
    }

    /**
     * Produces the transformed cell value for the current row and column.
     *
     * <p>Called by the framework only when {@link #needHandle()} returns {@code true}.
     * The return value is passed directly to the cell-value writer, so it must be a
     * type that Apache POI can handle: {@link String}, {@link Number},
     * {@link java.util.Date}, {@link Boolean}, or {@code null}.
     *
     * <p>The default implementation throws {@link UnsupportedOperationException} to
     * ensure that subclasses which opt in via {@code needHandle()} always provide a
     * concrete implementation.
     *
     * @param rowData context for the current row, giving access to the source data
     *                object, current cell and row indices, and raw POI primitives
     * @return the value to write into the current cell; must not be a complex object
     * @throws UnsupportedOperationException if this handler is activated but not implemented
     */
    default Object handler(ExcelRowData rowData) {
        throw new UnsupportedOperationException();
    }
}
