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

import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.io.IOException;

/**
 * Business-level extension of {@link SXSSFSheet} that correctly tracks the last written
 * row number across streaming flushes.
 *
 * <p>The standard {@link SXSSFSheet#getLastRowNum()} relies on in-memory row tracking and
 * returns {@code 0} after rows have been flushed to the temporary disk buffer, making it
 * unreliable for export pipelines that need to know the total number of rows written.
 * This class maintains its own {@code lastRowNum} counter that is updated on every call to
 * {@link #createRow(int)} and {@link #changeRowNum(SXSSFRow, int)}, so
 * {@link #getLastRowNum()} always reflects the highest row index written regardless of how
 * many rows have been flushed.</p>
 *
 * <p>Instances are created exclusively by {@link BusinessSXSSFWorkbook}.</p>
 *
 * @author dhsolo
 * @since 1.0
 * @see BusinessSXSSFWorkbook
 */
public class BusinessSXSSFSheet extends SXSSFSheet {

    /** The highest row index written to this sheet so far. */
    /** Highest row index seen; -1 while the sheet is empty (POI contract). */
    private int lastRowNum = -1;

    /**
     * Creates a streaming sheet backed by the given {@link XSSFSheet}.
     *
     * @param workbook the owning {@link SXSSFWorkbook}
     * @param xSheet   the underlying XSSF sheet that holds persistent sheet metadata
     * @throws IOException if the row-flush buffer cannot be initialised
     */
    public BusinessSXSSFSheet(SXSSFWorkbook workbook, XSSFSheet xSheet) throws IOException {
        super(workbook, xSheet);
    }

    /**
     * Creates a new row at the given row index and updates the internal last-row counter.
     *
     * @param rownum the zero-based row index
     * @return the newly created {@link SXSSFRow}
     */
    @Override
    public SXSSFRow createRow(int rownum) {
        SXSSFRow row = super.createRow(rownum);
        if(rownum > lastRowNum){
            lastRowNum = rownum;
        }
        return row;
    }

    /**
     * Moves {@code row} to {@code newRowNum} and updates the internal last-row counter if
     * the new index is higher than the current maximum.
     *
     * @param row       the row to move
     * @param newRowNum the target zero-based row index
     */
    @Override
    public void changeRowNum(SXSSFRow row, int newRowNum) {
        super.changeRowNum(row, newRowNum);
        if(newRowNum > lastRowNum){
            lastRowNum = newRowNum;
        }
    }

    /**
     * Returns the highest row index written to this sheet.
     *
     * <p>Unlike the base-class implementation, this value remains accurate after rows have
     * been flushed from the in-memory window to disk.</p>
     *
     * @return the zero-based index of the last row written
     */
    @Override
    public int getLastRowNum() {
        return lastRowNum;
    }
}
