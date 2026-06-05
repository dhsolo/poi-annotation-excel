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
package io.github.dhsolo.poi.excel.style;

import org.apache.poi.ss.usermodel.*;;

/**
 * Default implementation of {@link CellStyleManager}.
 *
 * <p>Creates and manages title, header, and data cell styles for Excel workbooks.
 * Styles are initialized with Song typeface and centered alignment with borders.
 *
 * @author dhsolo
 * @since 1.0
 */
public class DefaultCellStyleManager implements CellStyleManager {

    private final Workbook book;
    private CellStyle titleCellStyle;
    private CellStyle cellStyle;
    private CellStyle headerCellStyle;

    /**
     * Creates a style manager bound to the given workbook.
     * Call {@link #initDefaultStyles()} after construction to populate the three style roles.
     *
     * @param book the workbook on which cell styles will be created
     */
    public DefaultCellStyleManager(Workbook book) {
        this.book = book;
    }

    public void initDefaultStyles() {
        cellStyle = book.createCellStyle();
        Font font = book.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 12);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        headerCellStyle = book.createCellStyle();
        headerCellStyle.cloneStyleFrom(cellStyle);
        font = book.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        headerCellStyle.setFont(font);

        titleCellStyle = book.createCellStyle();
        titleCellStyle.cloneStyleFrom(cellStyle);
        font = book.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        titleCellStyle.setFont(font);
    }

    public CellStyle getTitleCellStyle() {
        return titleCellStyle;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public CellStyle getHeaderCellStyle() {
        return headerCellStyle;
    }

    public void setTitleCellStyle(CellStyle titleCellStyle) {
        this.titleCellStyle = titleCellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public void setHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
    }

    public CellStyle copyTitleStyle() {
        CellStyle style = book.createCellStyle();
        style.cloneStyleFrom(titleCellStyle);
        return style;
    }

    public CellStyle copyCellStyle() {
        CellStyle style = book.createCellStyle();
        style.cloneStyleFrom(cellStyle);
        return style;
    }

    public CellStyle copyHeaderStyle() {
        CellStyle style = book.createCellStyle();
        style.cloneStyleFrom(headerCellStyle);
        return style;
    }
}
