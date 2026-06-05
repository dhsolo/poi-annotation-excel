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

import io.github.dhsolo.poi.excel.render.ExcelTranslateHandler;
import io.github.dhsolo.poi.excel.validation.ExcelCustomValidate;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.Map;

/**
 * Read-only context object passed to callback hooks ({@link RowDataMapper},
 * {@link ExcelTranslateHandler}, and {@link ExcelCustomValidate}) during both Excel
 * export and import operations.
 *
 * <p>An {@code ExcelRowData} instance is scoped to a single cell: it captures the
 * position ({@link #currentRowNum()}, {@link #currentCellNum()}), the POI primitives
 * ({@link #getRow()}, {@link #getCell()}), the typed source/target data object
 * ({@link #getRowData()}), and — during import — the original raw cell values read
 * from the spreadsheet ({@link #getOriginalCellData()}).
 *
 * <p>Implementations of this interface are created and managed by the framework; callers
 * should not instantiate them directly.
 *
 * @param <T> the type of the domain object associated with the current row
 * @author dhsolo
 * @since 1.0
 */
public interface ExcelRowData<T> {

    /**
     * Returns the POI {@link Row} that is currently being processed.
     *
     * @return the current row; never {@code null} during a callback invocation
     */
    Row getRow();

    /**
     * Returns the POI {@link Cell} that is currently being processed.
     *
     * @return the current cell; never {@code null} during a callback invocation
     */
    Cell getCell();

    /**
     * Returns the domain object bound to the current row, typed as {@code T}.
     *
     * <p>During export, this is the source object from the data list.  During import,
     * this is the partially-populated target object being assembled from cell values.
     *
     * @return the row-level domain object; may be {@code null} if not yet populated
     */
    T getRowData();

    /**
     * Returns the domain object bound to the current row, cast to the specified type.
     * Convenience overload for use-cases where the generic parameter is not available
     * at the call site (e.g. in un-parameterised callback contexts).
     *
     * @param <U>    the desired target type
     * @param uClass the {@link Class} token for the target type; must not be {@code null}
     * @return the row-level domain object cast to {@code U}; may be {@code null}
     * @throws ClassCastException if the actual object cannot be cast to {@code U}
     */
    <U> U getRowData(Class<U> uClass);

    /**
     * Returns the value that the framework has resolved for the current cell.
     *
     * <p>During export, this is the field value read from the source data object before
     * any translation is applied.  During import, this is the converted cell value after
     * type coercion.
     *
     * @return the current cell value; may be {@code null}
     */
    Object currentValue();

    /**
     * Returns the zero-based column index of the cell currently being processed.
     *
     * @return the column index (&ge; 0)
     */
    int currentCellNum();

    /**
     * Returns the zero-based row index of the row currently being processed, counting
     * from the very first row of the sheet (including any title or header rows).
     *
     * @return the row index (&ge; 0)
     */
    int currentRowNum();

    /**
     * Returns the unmodified raw cell values read from the current Excel row, keyed by
     * the column header name, cast to the specified type.
     *
     * <p>This is primarily useful during import validation, where a callback needs access
     * to sibling columns that have not yet been mapped to the domain object.
     *
     * @param <U>    the desired return type
     * @param uClass the {@link Class} token for the target type; must not be {@code null}
     * @return the raw cell data map cast to {@code U}; may be {@code null}
     * @throws ClassCastException if the raw data cannot be cast to {@code U}
     */
    <U> U getOriginalCellData(Class<U> uClass);

    /**
     * Returns a map of all unmodified raw cell values for the current row, where each
     * key is the column header string and each value is the POI-extracted cell value
     * (a {@link String}, {@link Double}, {@link Boolean}, or {@link java.util.Date}).
     *
     * <p>Use this map during import callbacks to inspect columns other than the one
     * currently being processed, without triggering further type conversion.
     *
     * @return a non-null map from header name to raw cell value; values may be {@code null}
     *         for blank cells
     */
    Map<String, Object> getOriginalCellData();
}
