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

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Collections;
import java.util.List;

/**
 * Lightweight stub implementation of the Apache POI {@link Workbook} interface for
 * reading CSV files through the same export/import pipeline used for XLS and XLSX formats.
 *
 * <p>By implementing {@link Workbook}, this class allows the framework to treat a CSV
 * source as a single-sheet workbook without any special-casing in the pipeline. The
 * workbook always contains exactly one {@link CSVSheet} that is initialised from the
 * supplied {@link InputStream} during construction.</p>
 *
 * <p>Methods that have no meaningful equivalent in the CSV model (font management, cell
 * styles, print areas, named ranges, etc.) are no-ops or return sensible default values.
 * They are not annotated individually because their behaviour is self-evident from the
 * CSV context.</p>
 *
 * @author dh
 * @since 1.0
 * @see CSVSheet
 */
public class CSVWorkbook implements Workbook {

    private InputStream in;
    private Sheet sheet;

    /**
     * Constructs a {@code CSVWorkbook} by parsing the given input stream as a CSV file.
     *
     * <p>The input stream is consumed and parsed into an in-memory list of
     * {@link CSVRow}/{@link CSVCell} objects held by the single {@link CSVSheet}.
     * The character encoding is auto-detected; UTF-8 is used as the fallback.</p>
     *
     * @param in the CSV data stream; must not be {@code null}
     * @throws IOException if the stream cannot be read or CSV parsing fails
     */
    public CSVWorkbook(InputStream in) throws IOException {
        this.in = in;
        init();
    }

    /**
     * Initialises the single sheet by delegating to {@link CSVSheet}.
     *
     * @throws IOException if the underlying CSV stream cannot be parsed
     */
    private void init() throws IOException {
        this.sheet = new CSVSheet(in);
    }

    /** Returns an iterator over the single CSV sheet. */
    @Override public java.util.Iterator<Sheet> sheetIterator() { return java.util.Collections.<Sheet>singletonList(sheet).iterator(); }
    /** Returns {@code 0}; only one sheet exists. */
    @Override public int getActiveSheetIndex() { return 0; }
    /** Not supported for CSV format. */
    @Override public void setActiveSheet(int sheetIndex) {}
    /** Returns {@code 0}; only one sheet exists. */
    @Override public int getFirstVisibleTab() { return 0; }
    /** Not supported for CSV format. */
    @Override public void setFirstVisibleTab(int sheetIndex) {}
    /** Not supported for CSV format. */
    @Override public void setSheetOrder(String sheetname, int pos) {}
    /** Not supported for CSV format. */
    @Override public void setSelectedTab(int index) {}
    /** Not supported for CSV format. */
    @Override public void setSheetName(int sheet, String name) {}
    /** Returns an empty string; CSV sheets have no named tabs. */
    @Override public String getSheetName(int sheet) { return ""; }
    /** Returns {@code 0}; only one sheet exists. */
    @Override public int getSheetIndex(String name) { return 0; }
    /** Returns {@code 0}; only one sheet exists. */
    @Override public int getSheetIndex(Sheet sheet) { return 0; }
    /** Not supported for CSV format. */
    @Override public Sheet createSheet() { return null; }
    /** Not supported for CSV format. */
    @Override public Sheet createSheet(String sheetname) { return null; }
    /** Not supported for CSV format. */
    @Override public Sheet cloneSheet(int sheetNum) { return null; }
    /** Returns {@code 1}; a CSV workbook always contains exactly one sheet. */
    @Override public int getNumberOfSheets() { return 1; }

    /**
     * Returns the single CSV sheet regardless of {@code index}.
     *
     * @param index ignored; always returns the sole sheet
     * @return the {@link CSVSheet} constructed from the input stream
     */
    @Override public Sheet getSheetAt(int index) { return sheet; }

    /**
     * Returns the single CSV sheet regardless of {@code name}.
     *
     * @param name ignored; always returns the sole sheet
     * @return the {@link CSVSheet} constructed from the input stream
     */
    @Override public Sheet getSheet(String name) { return sheet; }
    /** Not supported for CSV format. */
    @Override public void removeSheetAt(int index) {}
    /** Not supported for CSV format. */
    @Override public Font createFont() { return null; }

    /** Not supported for CSV format. */
    @Override
    public Font findFont(boolean bold, short color, short fontHeight, String name,
                         boolean italic, boolean strikeout, short typeOffset, byte underline) {
        return null;
    }

    /** Not supported for CSV format. */
    @Override public int getNumberOfFonts() { return 0; }
    /** Not supported for CSV format. */
    @Override public int getNumberOfFontsAsInt() { return 0; }
    /** Not supported for CSV format. */
    @Override public Font getFontAt(int idx) { return null; }
    /** Not supported for CSV format. */
    @Override public CellStyle createCellStyle() { return null; }
    /** Not supported for CSV format. */
    @Override public int getNumCellStyles() { return 0; }
    /** Not supported for CSV format. */
    @Override public CellStyle getCellStyleAt(int idx) { return null; }
    /** Not supported for CSV format. */
    @Override public void write(OutputStream stream) throws IOException {}
    /** Not supported for CSV format. */
    @Override public void close() throws IOException {}
    /** Not supported for CSV format. */
    @Override public int getNumberOfNames() { return 0; }
    /** Not supported for CSV format. */
    @Override public Name getName(String name) { return null; }
    /** Not supported for CSV format. */
    @Override public List<? extends Name> getNames(String name) { return Collections.emptyList(); }
    /** Not supported for CSV format. */
    @Override public List<? extends Name> getAllNames() { return Collections.emptyList(); }
    /** Not supported for CSV format. */
    @Override public Name createName() { return null; }
    /** Not supported for CSV format. */
    @Override public void removeName(Name name) {}
    /** Not supported for CSV format. */
    @Override public void setPrintArea(int sheetIndex, String reference) {}
    /** Not supported for CSV format. */
    @Override public void setPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow) {}
    /** Not supported for CSV format. */
    @Override public String getPrintArea(int sheetIndex) { return ""; }
    /** Not supported for CSV format. */
    @Override public void removePrintArea(int sheetIndex) {}
    /** Not supported for CSV format. */
    @Override public Row.MissingCellPolicy getMissingCellPolicy() { return null; }
    /** Not supported for CSV format. */
    @Override public void setMissingCellPolicy(Row.MissingCellPolicy missingCellPolicy) {}
    /** Not supported for CSV format. */
    @Override public DataFormat createDataFormat() { return null; }
    /** Not supported for CSV format. */
    @Override public int addPicture(byte[] pictureData, int format) { return 0; }
    /** Not supported for CSV format. */
    @Override public List<? extends PictureData> getAllPictures() { return Collections.emptyList(); }
    /** Not supported for CSV format. */
    @Override public CreationHelper getCreationHelper() { return null; }
    /** Not supported for CSV format. */
    @Override public boolean isHidden() { return false; }
    /** Not supported for CSV format. */
    @Override public void setHidden(boolean hiddenFlag) {}
    /** Not supported for CSV format. */
    @Override public boolean isSheetHidden(int sheetIx) { return false; }
    /** Not supported for CSV format. */
    @Override public boolean isSheetVeryHidden(int sheetIx) { return false; }
    /** Not supported for CSV format. */
    @Override public void setSheetHidden(int sheetIx, boolean hidden) {}
    /** Not supported for CSV format. */
    @Override public SheetVisibility getSheetVisibility(int sheetIx) { return SheetVisibility.VISIBLE; }
    /** Not supported for CSV format. */
    @Override public void setSheetVisibility(int sheetIx, SheetVisibility visibility) {}
    /** Not supported for CSV format. */
    @Override public void addToolPack(UDFFinder toopack) {}
    /** Not supported for CSV format. */
    @Override public void setForceFormulaRecalculation(boolean value) {}
    /** Not supported for CSV format. */
    @Override public boolean getForceFormulaRecalculation() { return false; }
    /** Not supported for CSV format. */
    @Override public SpreadsheetVersion getSpreadsheetVersion() { return SpreadsheetVersion.EXCEL2007; }
    /** Not supported for CSV format. */
    @Override public int addOlePackage(byte[] oleData, String label, String fileName, String command) throws IOException { return 0; }
    /** Not supported for CSV format. */
    @Override public void setCellReferenceType(CellReferenceType cellReferenceType) {}
    /** Not supported for CSV format. */
    @Override public CellReferenceType getCellReferenceType() { return CellReferenceType.A1; }
    /** Not supported for CSV format. */
    @Override public org.apache.poi.ss.formula.EvaluationWorkbook createEvaluationWorkbook() { return null; }
    /** Not supported for CSV format. */
    @Override public int linkExternalWorkbook(String name, Workbook workbook) { return 0; }
}
