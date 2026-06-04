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
package io.github.dh.poi.excel.core;

import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Lightweight stub implementation of the Apache POI {@link Row} interface for a single
 * row of data parsed from a CSV file.
 *
 * <p>On construction the raw {@code String[]} produced by the CSV parser is converted into
 * a list of {@link CSVCell} objects so that the standard POI cell-access methods
 * ({@link #getCell(int)}, {@link #getFirstCellNum()}, {@link #getLastCellNum()}) work
 * correctly within the import pipeline.</p>
 *
 * <p>Methods that have no meaningful equivalent in the CSV model (row height, styles,
 * cell shifting, etc.) are no-ops or return safe default values.</p>
 *
 * @author dh
 * @since 1.0
 * @see CSVSheet
 * @see CSVCell
 */
public class CSVRow implements Row {

    /** The raw CSV values for this row, one element per column. */
    private String[] rowArray;

    /** Ordered list of cells created from {@link #rowArray}. */
    private List<Cell> cells = new ArrayList<>();

    /** The parent sheet that contains this row. */
    private Sheet sheet;

    /** Zero-based row index within the parent sheet. */
    private int index;

    /**
     * Constructs a {@code CSVRow} from the given CSV token array.
     *
     * @param rowArray the parsed CSV tokens for this row; must not be {@code null}
     * @param sheet    the parent {@link CSVSheet}
     * @param index    zero-based row index within the sheet
     */
    public CSVRow(String[] rowArray, Sheet sheet, int index) {
        this.rowArray = rowArray;
        this.sheet = sheet;
        this.index = index;
        init();
    }

    /**
     * Wraps each token in {@link #rowArray} as a {@link CSVCell} and adds it to the
     * {@link #cells} list.
     */
    private void init() {
        for (int i = 0; i < rowArray.length; i++) {
            cells.add(new CSVCell(rowArray[i], this, i));
        }
    }

    /** Not supported for CSV format. */
    @Override
    public Cell createCell(int column) {
        return null;
    }

    /** Not supported for CSV format. */
    @Override
    public Cell createCell(int column, CellType type) {
        return null;
    }

    /** Not supported for CSV format. */
    @Override
    public void shiftCellsLeft(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {}

    /** Not supported for CSV format. */
    @Override
    public void shiftCellsRight(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {}

    /** Not supported for CSV format. */
    @Override
    public int getOutlineLevel() { return 0; }

    /** Not supported for CSV format. */
    @Override
    public void removeCell(Cell cell) {

    }

    /** Not supported for CSV format. */
    @Override
    public void setRowNum(int rowNum) {

    }

    /**
     * Returns the zero-based row index of this row within the parent sheet.
     *
     * @return the row index supplied at construction time
     */
    @Override
    public int getRowNum() {
        return index;
    }

    /**
     * Returns the cell at the given zero-based column index.
     *
     * @param cellnum the zero-based column index
     * @return the {@link CSVCell} at {@code cellnum}
     * @throws IndexOutOfBoundsException if {@code cellnum} is out of range
     */
    @Override
    public Cell getCell(int cellnum) {
        return cells.get(cellnum);
    }

    /** Not supported for CSV format. */
    @Override
    public Cell getCell(int cellnum, MissingCellPolicy policy) {
        return null;
    }

    /**
     * Returns {@code 0}; CSV rows always start from the first column.
     *
     * @return {@code 0}
     */
    @Override
    public short getFirstCellNum() {
        return 0;
    }

    /**
     * Returns the zero-based index of the last cell in this row.
     *
     * @return {@code cells.size() - 1}
     */
    @Override
    public short getLastCellNum() {
        return (short) (cells.size() - 1);
    }

    /** Not supported for CSV format. */
    @Override
    public int getPhysicalNumberOfCells() {
        return 0;
    }

    /** Not supported for CSV format. */
    @Override
    public void setHeight(short height) {

    }

    /** Not supported for CSV format. */
    @Override
    public void setZeroHeight(boolean zHeight) {

    }

    /** Not supported for CSV format. */
    @Override
    public boolean getZeroHeight() {
        return false;
    }

    /** Not supported for CSV format. */
    @Override
    public void setHeightInPoints(float height) {

    }

    /** Not supported for CSV format. */
    @Override
    public short getHeight() {
        return 0;
    }

    /** Not supported for CSV format. */
    @Override
    public float getHeightInPoints() {
        return 0;
    }

    /** Not supported for CSV format. */
    @Override
    public boolean isFormatted() {
        return false;
    }

    /** Not supported for CSV format. */
    @Override
    public CellStyle getRowStyle() {
        return null;
    }

    /** Not supported for CSV format. */
    @Override
    public void setRowStyle(CellStyle style) {

    }

    /** Not supported for CSV format. */
    @Override
    public Iterator<Cell> cellIterator() {
        return null;
    }

    /**
     * Returns the parent sheet of this row.
     *
     * @return the {@link CSVSheet} that contains this row
     */
    @Override
    public Sheet getSheet() {
        return sheet;
    }

    /** Not supported for CSV format. */
    @Override
    public Iterator<Cell> iterator() {
        return null;
    }
}
