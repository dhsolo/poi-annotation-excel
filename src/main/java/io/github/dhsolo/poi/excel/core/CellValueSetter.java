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

import io.github.dhsolo.poi.excel.model.ExcelRowData;
import io.github.dhsolo.poi.excel.model.RowDataMapper;
import io.github.dhsolo.poi.excel.render.ExcelTranslateHandler;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Strategy interface for writing a value into a POI {@link Cell} and for constructing
 * the {@link ExcelRowData} context objects used by callback hooks during export.
 *
 * <p>Follows Spring's {@code ResourceLoader} pattern: render-layer components depend only
 * on this narrow interface rather than on the full {@code ExcelCreator}, keeping the
 * module boundary clean and the contract easily mockable in tests.
 *
 * <p>The framework supplies a single implementation that handles all standard POI-supported
 * cell-value types ({@link String}, {@link Number}, {@link java.util.Date}, {@link Boolean}).
 * Custom type coercion can be achieved by wrapping or replacing this implementation.
 *
 * @author dh
 * @since 1.0
 */
public interface CellValueSetter {

    /**
     * Writes {@code value} into the given POI {@link Cell}, applying the appropriate
     * POI setter method based on the runtime type of {@code value}.
     *
     * <p>Type mapping:
     * <ul>
     *   <li>{@link Boolean} — {@code cell.setCellValue(boolean)}</li>
     *   <li>{@link Number} — {@code cell.setCellValue(double)}</li>
     *   <li>{@link java.util.Date} / {@link java.util.Calendar} — date-formatted numeric cell</li>
     *   <li>Everything else (including {@link String}) — {@code cell.setCellValue(String)}</li>
     *   <li>{@code null} — the cell is left blank ({@code cell.setBlank()})</li>
     * </ul>
     *
     * @param cell  the target POI cell; must not be {@code null}
     * @param value the value to write; may be {@code null} to produce a blank cell
     */
    void setCellValue(Cell cell, Object value);

    /**
     * Writes {@code value} into the given cell, treating temporal values as typed, date-formatted
     * cells. When {@code value} is a {@link java.util.Date}, {@link java.util.Calendar},
     * {@link java.time.LocalDate}, or {@link java.time.LocalDateTime}, it is written with its
     * native type and a cell style using {@code datePattern} (or a type-based default when
     * {@code datePattern} is {@code null}); all other values delegate to
     * {@link #setCellValue(Cell, Object)}.
     *
     * @param cell        the target POI cell; must not be {@code null}
     * @param value       the value to write; may be {@code null}
     * @param datePattern the date format pattern for temporal values, or {@code null} for a default
     */
    default void setCellValue(Cell cell, Object value, String datePattern) {
        setCellValue(cell, value);
    }

    /**
     * Builds an {@link ExcelRowData} context object that encapsulates the current position
     * and data for use by {@link RowDataMapper} and {@link ExcelTranslateHandler} callbacks.
     *
     * @param obj      the source data object for the current row
     * @param field    the name of the field being written in the current column
     * @param colIndex the zero-based column index of the cell being written
     * @param rowIndex the zero-based row index of the cell being written
     * @return a fully populated {@link ExcelRowData} instance; never {@code null}
     */
    ExcelRowData createExcelRowData(Object obj, String field, int colIndex, int rowIndex);
}
