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
package io.github.dhsolo.poi.excel.model;

import io.github.dhsolo.poi.excel.ExcelCreator;
import io.github.dhsolo.poi.excel.core.CellValueSetter;

/**
 * Functional callback interface for supplying a custom cell value during Excel export.
 *
 * <p>Modelled after Spring's {@code RowMapper} pattern. Register a {@code RowDataMapper}
 * with {@code ExcelCreator} to override the default field-reflection approach for one or
 * more columns.  The framework invokes {@link #mapRow(ExcelRowData)} once per cell, passing
 * a rich context object that exposes the source data object, the current row and column
 * indices, and the underlying POI primitives.
 *
 * <p>This is a {@link FunctionalInterface}, so implementations can be supplied as lambda
 * expressions or method references:
 * <pre>{@code
 * creator.setRowDataMapper((ExcelRowData<Order> ctx) -> {
 *     if (ctx.currentCellNum() == 3) {
 *         return ctx.getRowData().getPrice().toPlainString();
 *     }
 *     return null; // fall back to default for other columns
 * });
 * }</pre>
 *
 * @param <T> the type of the domain object bound to each row
 * @author dhsolo
 * @since 1.0
 * @see ExcelCreator
 */
@FunctionalInterface
public interface RowDataMapper<T> {

    /**
     * Produces the value to write into the current cell, or {@code null} to fall back to
     * the framework's default field-reflection behaviour for that column.
     *
     * <p>The returned value is passed directly to the {@link CellValueSetter} and must be
     * a POI-compatible type: {@link String}, {@link Number}, {@link java.util.Date},
     * {@link Boolean}, or {@code null}.  Returning a complex object results in its
     * {@code toString()} representation being written as a string cell value.
     *
     * @param rowData contextual information for the current cell, including the typed
     *                source data object ({@link ExcelRowData#getRowData()}), the current
     *                column index ({@link ExcelRowData#currentCellNum()}), the row index
     *                ({@link ExcelRowData#currentRowNum()}), and direct access to the POI
     *                {@link org.apache.poi.ss.usermodel.Cell} and {@link org.apache.poi.ss.usermodel.Row}
     * @return the value to write, or {@code null} to delegate to the default handling
     */
    Object mapRow(ExcelRowData<T> rowData);
}
