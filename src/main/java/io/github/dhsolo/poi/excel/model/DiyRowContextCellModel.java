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

import io.github.dhsolo.poi.excel.style.CellStyleEnum;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Describes a single custom ("DIY") cell that the framework should inject into a
 * non-data row of the exported sheet.
 *
 * <p>Custom rows sit outside the standard title/header/data row structure.  Each
 * {@code DiyRowContextCellModel} instance defines one cell (or a merged range of cells)
 * within such a row, specifying the text to display, the column span, and the visual
 * style to apply.
 *
 * <p>The {@link #isAfterTitle()} flag controls placement:
 * <ul>
 *   <li>{@code true} — the row is inserted after the title row but before the column
 *       header row, typically used for sub-titles, date stamps, or remarks.</li>
 *   <li>{@code false} — the row is inserted at the default custom-row position defined
 *       by the surrounding context.</li>
 * </ul>
 *
 * <p>If the cell should span multiple columns, set {@link #setStartColumn(int)} and
 * {@link #setEndColumn(int)} to the first and last (inclusive) column indices; the
 * framework will merge the cells automatically.
 *
 * @author dh
 * @since 1.0
 */
public class DiyRowContextCellModel {

    /**
     * When {@code true}, the custom row containing this cell is placed after the title
     * row and before the column header row.  When {@code false}, placement follows the
     * default custom-row ordering.
     */
    private boolean isAfterTitle;

    /**
     * Zero-based sequential index of this cell within its containing custom row.
     * Used by the framework to order cells when multiple {@code DiyRowContextCellModel}
     * instances belong to the same row.
     */
    private int diyRowContextCellNum;

    /**
     * The text string to write into this cell.  May be {@code null} to leave the cell empty.
     */
    private String value;

    /**
     * Zero-based index of the first column that this cell occupies (inclusive).
     * If no merge is needed, set this equal to {@link #endColumn}.
     */
    private int startColumn;

    /**
     * Zero-based index of the last column that this cell occupies (inclusive).
     * If no merge is needed, set this equal to {@link #startColumn}.
     */
    private int endColumn;

    /**
     * Logical style role for this cell.  Defaults to {@link CellStyleEnum#defaultStyle}.
     * Set to {@link CellStyleEnum#customStyle} and also call
     * {@link #setCustomCellStyle(CellStyle)} when a fully custom POI style is required.
     */
    private CellStyleEnum cellStyleEnum = CellStyleEnum.defaultStyle;

    /**
     * A caller-supplied POI {@link CellStyle} applied when {@link #cellStyleEnum} is
     * {@link CellStyleEnum#customStyle}.  Ignored for all other style roles.
     */
    private CellStyle customCellStyle;

    /**
     * Returns the zero-based sequential index of this cell within its custom row.
     *
     * @return the cell index within the custom row
     */
    public int getDiyRowContextCellNum() {
        return diyRowContextCellNum;
    }

    /**
     * Sets the zero-based sequential index of this cell within its custom row.
     * The framework uses this value to order cells when building the row.
     *
     * @param diyRowContextCellNum the cell index; must be unique within the row
     */
    public void setDiyRowContextCellNum(int diyRowContextCellNum) {
        this.diyRowContextCellNum = diyRowContextCellNum;
    }

    /**
     * Returns the zero-based first column index (inclusive) of this cell's merge range.
     *
     * @return the starting column index
     */
    public int getStartColumn() {
        return startColumn;
    }

    /**
     * Sets the zero-based first column index (inclusive) for this cell.
     * When this value equals {@link #getEndColumn()}, no merge is performed.
     *
     * @param startColumn the starting column index; must be &ge; 0
     */
    public void setStartColumn(int startColumn) {
        this.startColumn = startColumn;
    }

    /**
     * Returns the zero-based last column index (inclusive) of this cell's merge range.
     *
     * @return the ending column index
     */
    public int getEndColumn() {
        return endColumn;
    }

    /**
     * Sets the zero-based last column index (inclusive) for this cell.
     * When this value equals {@link #getStartColumn()}, no merge is performed.
     *
     * @param endColumn the ending column index; must be &ge; {@code startColumn}
     */
    public void setEndColumn(int endColumn) {
        this.endColumn = endColumn;
    }

    /**
     * Returns the text content to be written into this cell.
     *
     * @return the cell text, or {@code null} if the cell is to be left empty
     */
    public String getValue() {
        return value;
    }

    /**
     * Sets the text content to write into this cell.
     *
     * @param value the cell text; {@code null} results in an empty cell
     */
    public void setValue(String value) {
        this.value = value;
    }

    /**
     * Returns the logical style role assigned to this cell.
     *
     * @return the {@link CellStyleEnum} constant; never {@code null}
     */
    public CellStyleEnum getCellStyleEnum() {
        return cellStyleEnum;
    }

    /**
     * Sets the logical style role for this cell.
     * To supply a fully custom POI {@link CellStyle}, set this to
     * {@link CellStyleEnum#customStyle} and call {@link #setCustomCellStyle(CellStyle)}.
     *
     * @param cellStyleEnum the style role to apply; must not be {@code null}
     */
    public void setCellStyleEnum(CellStyleEnum cellStyleEnum) {
        this.cellStyleEnum = cellStyleEnum;
    }

    /**
     * Returns the caller-supplied POI {@link CellStyle} used when the style role is
     * {@link CellStyleEnum#customStyle}.
     *
     * @return the custom cell style, or {@code null} if none has been set
     */
    public CellStyle getCustomCellStyle() {
        return customCellStyle;
    }

    /**
     * Sets a caller-supplied POI {@link CellStyle} to use for this cell.
     * This value is only applied when {@link #getCellStyleEnum()} returns
     * {@link CellStyleEnum#customStyle}.
     *
     * @param customCellStyle the POI cell style to apply; may be {@code null}
     */
    public void setCustomCellStyle(CellStyle customCellStyle) {
        this.customCellStyle = customCellStyle;
    }

    /**
     * Returns whether the custom row containing this cell is positioned after the
     * title row and before the column header row.
     *
     * @return {@code true} if the row appears between the title row and the header row;
     *         {@code false} for the default placement
     */
    public boolean isAfterTitle() {
        return isAfterTitle;
    }

    /**
     * Controls whether the custom row containing this cell is inserted between the
     * title row and the column header row.
     *
     * @param isAfterTitle {@code true} to place the row after the title and before the
     *                     header; {@code false} for the default placement
     */
    public void setAfterTitle(boolean isAfterTitle) {
        this.isAfterTitle = isAfterTitle;
    }
}
