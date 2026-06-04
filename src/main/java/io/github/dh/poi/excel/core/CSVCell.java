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

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

/**
 * Lightweight stub implementation of the Apache POI {@link Cell} interface for a single
 * field value parsed from a CSV file.
 *
 * <p>A CSV cell holds a single {@code String} value. All CSV data is treated as plain text,
 * so {@link #getCellType()} always returns {@link CellType#STRING}. Numeric access via
 * {@link #getNumericCellValue()} attempts a {@link Double#parseDouble(String)} conversion
 * and returns {@code 0} on failure. Rich-text access wraps the raw string in an
 * {@link XSSFRichTextString}.</p>
 *
 * <p>Methods that have no CSV equivalent (formulas, dates, styles, hyperlinks, etc.) are
 * no-ops or return safe default values, enabling the import pipeline to call any
 * {@link Cell} method without special-casing the CSV code path.</p>
 *
 * @author dh
 * @since 1.0
 * @see CSVRow
 */
public class CSVCell implements Cell {

    /** The raw CSV field value for this cell. */
    private String value;
    /** The parent row that contains this cell. */
    private Row row;
    /** Zero-based column index within the parent row. */
    private Integer index;

    /**
     * Constructs a {@code CSVCell} with the given value, parent row, and column index.
     *
     * @param value the raw CSV field value; may be empty but not {@code null}
     * @param row   the parent {@link CSVRow}
     * @param index zero-based column index within the row
     */
    public CSVCell(String value, Row row, int index) {
        this.value = value;
        this.row = row;
        this.index = index;
    }

    /**
     * Returns the zero-based column index of this cell.
     *
     * @return the column index supplied at construction time
     */
    @Override
    public int getColumnIndex() {
        return index;
    }

    /**
     * Returns the zero-based row index of this cell's parent row.
     *
     * @return the row index from the parent {@link Row}
     */
    @Override
    public int getRowIndex() {
        return row.getRowNum();
    }

    /**
     * Returns the parent sheet of this cell.
     *
     * @return the {@link CSVSheet} that contains the parent row
     */
    @Override
    public Sheet getSheet() {
        return row.getSheet();
    }

    /**
     * Returns the parent row of this cell.
     *
     * @return the {@link CSVRow} that contains this cell
     */
    @Override
    public Row getRow() {
        return row;
    }

    /**
     * Returns {@link CellType#STRING}; all CSV cell values are treated as plain text.
     *
     * @return {@link CellType#STRING}
     */
    @Override
    public CellType getCellType() {
        return CellType.STRING;
    }

    /**
     * Returns {@link CellType#STRING}; CSV cells do not contain formulas.
     *
     * @return {@link CellType#STRING}
     */
    @Override
    public CellType getCachedFormulaResultType() {
        return CellType.STRING;
    }

    /** Not supported for CSV format. */
    @Override
    public void setCellValue(double value) {}

    /** Not supported for CSV format. */
    @Override
    public void setCellValue(Date value) {}

    /** Not supported for CSV format. */
    @Override
    public void setCellValue(LocalDateTime value) {}

    /** Not supported for CSV format. */
    @Override
    public void setCellValue(Calendar value) {}

    /** Not supported for CSV format. */
    @Override
    public void setCellValue(RichTextString value) {}

    /**
     * Sets the raw string value of this cell.
     *
     * @param value the new cell value; may be {@code null}, which will store {@code null}
     */
    @Override
    public void setCellValue(String value) {
        this.value = value;
    }

    /** Not supported for CSV format. */
    @Override
    public void setCellFormula(String formula) throws FormulaParseException {}

    /** Not supported for CSV format. */
    @Override
    public void removeFormula() {}

    /**
     * Clears the cell value by setting it to an empty string.
     */
    @Override
    public void setBlank() { this.value = ""; }

    /** Not supported for CSV format. */
    @Override
    public void setCellType(CellType cellType) {}

    /**
     * Returns an empty string; CSV cells do not contain formulas.
     *
     * @return an empty string
     */
    @Override
    public String getCellFormula() {
        return "";
    }

    /**
     * Attempts to parse the cell's string value as a {@code double}.
     *
     * @return the parsed numeric value, or {@code 0} if the value is not a valid number
     */
    @Override
    public double getNumericCellValue() {
        try {
            return Double.parseDouble(value);
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    /** Not supported for CSV format. */
    @Override
    public Date getDateCellValue() {
        return null;
    }

    /** Not supported for CSV format. */
    @Override
    public LocalDateTime getLocalDateTimeCellValue() {
        return null;
    }

    /**
     * Returns the cell value wrapped in an {@link XSSFRichTextString}.
     *
     * @return a new {@link XSSFRichTextString} containing the raw CSV field value
     */
    @Override
    public RichTextString getRichStringCellValue() {
        return new XSSFRichTextString(value);
    }

    /**
     * Returns the raw string value of this cell as read from the CSV file.
     *
     * @return the cell's string value; may be empty but not {@code null} after construction
     */
    @Override
    public String getStringCellValue() {
        return value;
    }

    /** Not supported for CSV format. */
    @Override
    public void setCellValue(boolean value) {}

    /** Not supported for CSV format. */
    @Override
    public void setCellErrorValue(byte value) {}

    /** Not supported for CSV format. */
    @Override
    public boolean getBooleanCellValue() {
        return false;
    }

    /** Not supported for CSV format. */
    @Override
    public byte getErrorCellValue() {
        return 0;
    }

    /** Not supported for CSV format. */
    @Override
    public void setCellStyle(CellStyle style) {}

    /** Not supported for CSV format. */
    @Override
    public CellStyle getCellStyle() {
        return null;
    }

    /** Not supported for CSV format. */
    @Override
    public void setAsActiveCell() {}

    /**
     * Returns the {@link CellAddress} of this cell derived from its row and column indices.
     *
     * @return a new {@link CellAddress} for this cell's position
     */
    @Override
    public CellAddress getAddress() {
        return new CellAddress(getRowIndex(), getColumnIndex());
    }

    /** Not supported for CSV format. */
    @Override
    public void setCellComment(Comment comment) {}

    /** Not supported for CSV format. */
    @Override
    public Comment getCellComment() {
        return null;
    }

    /** Not supported for CSV format. */
    @Override
    public void removeCellComment() {}

    /** Not supported for CSV format. */
    @Override
    public Hyperlink getHyperlink() {
        return null;
    }

    /** Not supported for CSV format. */
    @Override
    public void setHyperlink(Hyperlink link) {}

    /** Not supported for CSV format. */
    @Override
    public void removeHyperlink() {}

    /** Not supported for CSV format. */
    @Override
    public CellRangeAddress getArrayFormulaRange() {
        return null;
    }

    /** Not supported for CSV format. */
    @Override
    public boolean isPartOfArrayFormulaGroup() {
        return false;
    }
}
