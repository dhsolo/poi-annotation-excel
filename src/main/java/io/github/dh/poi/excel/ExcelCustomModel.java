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
package io.github.dh.poi.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * Mutable context object that carries the state of the cell currently being processed
 * during an Excel import operation.
 *
 * <p>An {@code ExcelCustomModel} is populated by the import engine for each cell in every
 * data row and passed to custom cell-processing hooks (such as per-column value converters
 * or validation callbacks).  It aggregates positional information (row and column indices),
 * the POI primitives ({@link Row}, {@link Cell}), the converted cell value, and the column
 * definition ({@link ExcelModel}) so that a callback has everything it needs without holding
 * a reference to the full import context.
 *
 * <p>Instances of this class are reused across cells within the same import session; do not
 * retain a reference beyond the scope of a single callback invocation.
 *
 * @author dh
 * @since 1.0
 */
public class ExcelCustomModel {

    /**
     * Zero-based index of the column currently being processed, matching the physical
     * column position in the Excel sheet.
     */
    private int currentCellNum;

    /**
     * Zero-based index of the row currently being processed, counting from the first
     * row of the sheet including any title or header rows.
     */
    private int currentRowNum;

    /**
     * The POI {@link Row} object for the row currently being processed, providing access
     * to all cells in the row and to row-level formatting metadata.
     */
    private Row row;

    /**
     * The POI {@link Cell} object for the cell currently being processed, providing
     * direct access to the underlying cell type and raw value.
     */
    private Cell cell;

    /**
     * The converted cell value after the import engine has applied type coercion based
     * on the target field type declared in the model class.  May be {@code null} for
     * blank cells or cells that could not be converted.
     */
    private Object currentValue;

    /**
     * The column definition that describes how this cell maps to the target model field,
     * including the field name, expected type, and any validation rules declared via
     * {@code @ExcelColumn}.
     */
    private ExcelModel excelModel;

    /**
     * Returns the converted value of the cell currently being processed.
     *
     * @return the cell value after type coercion, or {@code null} for blank or unconvertible cells
     */
    public Object getCurrentValue() {
        return currentValue;
    }

    /**
     * Sets the converted cell value.  Called by the import engine after type coercion;
     * may also be overridden by a custom value-converter callback.
     *
     * @param currentValue the coerced cell value; may be {@code null}
     */
    public void setCurrentValue(Object currentValue) {
        this.currentValue = currentValue;
    }

    /**
     * Returns the POI {@link Row} for the row currently being processed.
     *
     * @return the current row; never {@code null} during a callback invocation
     */
    public Row getRow() {
        return row;
    }

    /**
     * Sets the POI {@link Row} for the current import iteration.
     * Called by the import engine before invoking any cell-level callbacks.
     *
     * @param row the current POI row; must not be {@code null}
     */
    public void setRow(Row row) {
        this.row = row;
    }

    /**
     * Returns the POI {@link Cell} for the cell currently being processed.
     *
     * @return the current cell; never {@code null} during a callback invocation
     */
    public Cell getCell() {
        return cell;
    }

    /**
     * Sets the POI {@link Cell} for the current import iteration.
     * Called by the import engine as it advances through each column of a row.
     *
     * @param cell the current POI cell; must not be {@code null}
     */
    public void setCell(Cell cell) {
        this.cell = cell;
    }

    /**
     * Returns the zero-based column index of the cell currently being processed.
     *
     * @return the column index (&ge; 0)
     */
    public int getCurrentCellNum() {
        return currentCellNum;
    }

    /**
     * Sets the zero-based column index for the current import iteration.
     *
     * @param currentCellNum the column index; must be &ge; 0
     */
    public void setCurrentCellNum(int currentCellNum) {
        this.currentCellNum = currentCellNum;
    }

    /**
     * Returns the zero-based row index of the row currently being processed.
     *
     * @return the row index (&ge; 0)
     */
    public int getCurrentRowNum() {
        return currentRowNum;
    }

    /**
     * Sets the zero-based row index for the current import iteration.
     *
     * @param currentRowNum the row index; must be &ge; 0
     */
    public void setCurrentRowNum(int currentRowNum) {
        this.currentRowNum = currentRowNum;
    }

    /**
     * Returns the column definition ({@link ExcelModel}) that describes how the current
     * cell maps to its target field in the model class.
     *
     * @return the column definition; never {@code null} during a callback invocation
     */
    public ExcelModel getExcelModel() {
        return excelModel;
    }

    /**
     * Sets the column definition for the current import iteration.
     * Called by the import engine each time it advances to the next column.
     *
     * @param excelModel the column definition; must not be {@code null}
     */
    public void setExcelModel(ExcelModel excelModel) {
        this.excelModel = excelModel;
    }
}
