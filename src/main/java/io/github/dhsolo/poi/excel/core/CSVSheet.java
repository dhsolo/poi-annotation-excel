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

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.mozilla.universalchardet.UniversalDetector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Lightweight stub implementation of the Apache POI {@link Sheet} interface for reading
 * CSV files through the same import pipeline used for XLS and XLSX formats.
 *
 * <p>On construction the entire CSV stream is read into memory, the character encoding is
 * auto-detected via Mozilla's {@code UniversalDetector} (falling back to UTF-8), and every
 * CSV row is parsed into a {@link CSVRow} containing {@link CSVCell} objects. The resulting
 * list of rows is then accessible through the standard POI {@link Sheet} API
 * ({@link #getRow(int)}, {@link #getLastRowNum()}, etc.), letting the import pipeline
 * iterate over a CSV file in exactly the same way it processes spreadsheet files.</p>
 *
 * <p>Methods that have no meaningful equivalent in the CSV model (gridlines, print setup,
 * merged regions, freeze panes, etc.) are no-ops or return safe default values.</p>
 *
 * @author dhsolo
 * @since 1.0
 * @see CSVWorkbook
 * @see CSVRow
 */
public class CSVSheet implements Sheet {

    private static final Logger log = LoggerFactory.getLogger(CSVSheet.class);

    private InputStream in;

    /** In-memory list of all rows parsed from the CSV input. */
    private List<Row> rows = new ArrayList<>();

    /**
     * Constructs a {@code CSVSheet} by parsing the given CSV input stream.
     *
     * <p>The stream is fully consumed during construction. The character encoding is
     * auto-detected; UTF-8 is used as the fallback.</p>
     *
     * @param in the CSV data stream; must not be {@code null}
     * @throws IOException if the stream cannot be read or CSV parsing fails
     */
    public CSVSheet(InputStream in) throws IOException {
        this.in = in;
        init();
    }

    /**
     * Reads the raw bytes from the input stream, detects the character encoding, parses
     * all CSV rows, and populates the {@link #rows} list with {@link CSVRow} instances.
     *
     * @throws IOException if the stream cannot be read or the CSV is malformed
     */
    private void init() throws IOException {
        byte[] b = in.readAllBytes();
        UniversalDetector detector = new UniversalDetector(null);
        detector.handleData(b,0,b.length);
        detector.dataEnd();
        String charset = detector.getDetectedCharset();
        log.info("CSV charset:{},read",charset);
        in = new ByteArrayInputStream(b);
        List<String[]> rowsList;
        try (CSVReader csvReader = new CSVReader(new InputStreamReader(in, charset != null ? charset : "UTF-8"))) {
            rowsList = csvReader.readAll();
        } catch (CsvException e) {
            throw new IOException("CSV parsing failed", e);
        }
        for(int i = 0 ; i < rowsList.size() ; i++){
            String [] rowArray = rowsList.get(i);
            Row row = new CSVRow(rowArray,this,i);
            rows.add(row);
        }
    }

    /** Not supported for CSV format. */
    @Override
    public Row createRow(int rownum) {
        return null;
    }

    /** Not supported for CSV format. */
    @Override
    public void removeRow(Row row) {

    }

    /**
     * Returns the row at the given zero-based index.
     *
     * @param rownum the zero-based row index
     * @return the row at {@code rownum}, or {@code null} when out of range (POI contract)
     */
    @Override
    public Row getRow(int rownum) {
        if (rownum < 0 || rownum >= rows.size()) {
            return null;
        }
        return rows.get(rownum);
    }

    @Override
    public int getPhysicalNumberOfRows() {
        return rows.size();
    }

    /**
     * Returns {@code 0}; CSV sheets always start from the first row.
     *
     * @return {@code 0}
     */
    @Override
    public int getFirstRowNum() {
        return 0;
    }

    /**
     * Returns the zero-based index of the last row in this sheet.
     *
     * @return {@code rows.size() - 1}, or {@code -1} if the sheet is empty
     */
    @Override
    public int getLastRowNum() {
        return rows.size()-1;
    }

    @Override
    public void setColumnHidden(int columnIndex, boolean hidden) {

    }

    @Override
    public boolean isColumnHidden(int columnIndex) {
        return false;
    }

    @Override
    public void setRightToLeft(boolean value) {

    }

    @Override
    public boolean isRightToLeft() {
        return false;
    }

    @Override
    public void setColumnWidth(int columnIndex, int width) {

    }

    @Override
    public int getColumnWidth(int columnIndex) {
        return 0;
    }

    @Override
    public void setDefaultColumnWidth(int width) {

    }

    @Override
    public int getDefaultColumnWidth() {
        return 0;
    }

    @Override
    public short getDefaultRowHeight() {
        return 0;
    }

    @Override
    public float getDefaultRowHeightInPoints() {
        return 0;
    }

    @Override
    public void setDefaultRowHeight(short height) {

    }

    @Override
    public void setDefaultRowHeightInPoints(float height) {

    }

    @Override
    public CellStyle getColumnStyle(int column) {
        return null;
    }

    @Override
    public int addMergedRegion(CellRangeAddress region) {
        return 0;
    }

    @Override
    public void setVerticallyCenter(boolean value) {

    }

    @Override
    public void setHorizontallyCenter(boolean value) {

    }

    @Override
    public boolean getHorizontallyCenter() {
        return false;
    }

    @Override
    public boolean getVerticallyCenter() {
        return false;
    }

    @Override
    public void removeMergedRegion(int index) {

    }

    @Override
    public int getNumMergedRegions() {
        return 0;
    }

    @Override
    public CellRangeAddress getMergedRegion(int index) {
        return null;
    }

    @Override
    public Iterator<Row> rowIterator() {
        return rows.iterator();
    }

    @Override
    public void setForceFormulaRecalculation(boolean value) {

    }

    @Override
    public boolean getForceFormulaRecalculation() {
        return false;
    }

    @Override
    public void setAutobreaks(boolean value) {

    }

    @Override
    public void setDisplayGuts(boolean value) {

    }

    @Override
    public void setDisplayZeros(boolean value) {

    }

    @Override
    public boolean isDisplayZeros() {
        return false;
    }

    @Override
    public void setFitToPage(boolean value) {

    }

    @Override
    public void setRowSumsBelow(boolean value) {

    }

    @Override
    public void setRowSumsRight(boolean value) {

    }

    @Override
    public boolean getAutobreaks() {
        return false;
    }

    @Override
    public boolean getDisplayGuts() {
        return false;
    }

    @Override
    public boolean getFitToPage() {
        return false;
    }

    @Override
    public boolean getRowSumsBelow() {
        return false;
    }

    @Override
    public boolean getRowSumsRight() {
        return false;
    }

    @Override
    public boolean isPrintGridlines() {
        return false;
    }

    @Override
    public void setPrintGridlines(boolean show) {

    }

    @Override
    public PrintSetup getPrintSetup() {
        return null;
    }

    @Override
    public Header getHeader() {
        return null;
    }

    @Override
    public Footer getFooter() {
        return null;
    }

    @Override
    public void setSelected(boolean value) {

    }

    @Override
    public double getMargin(short margin) {
        return 0;
    }

    @Override
    public void setMargin(short margin, double size) {

    }

    @Override
    public boolean getProtect() {
        return false;
    }

    @Override
    public void protectSheet(String password) {

    }

    @Override
    public boolean getScenarioProtect() {
        return false;
    }

    @Override
    public void setZoom(int scale) {

    }

    @Override
    public short getTopRow() {
        return 0;
    }

    @Override
    public short getLeftCol() {
        return 0;
    }

    @Override
    public void showInPane(int toprow, int leftcol) {

    }

    @Override
    public void shiftRows(int startRow, int endRow, int n) {

    }

    @Override
    public void shiftRows(int startRow, int endRow, int n, boolean copyRowHeight, boolean resetOriginalRowHeight) {

    }

    @Override
    public void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow) {

    }

    @Override
    public void createFreezePane(int colSplit, int rowSplit) {

    }

    @Override
    public void createSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, int activePane) {

    }

    @Override
    public void setDisplayGridlines(boolean show) {

    }

    @Override
    public boolean isDisplayGridlines() {
        return false;
    }

    @Override
    public void setDisplayFormulas(boolean show) {

    }

    @Override
    public boolean isDisplayFormulas() {
        return false;
    }

    @Override
    public void setDisplayRowColHeadings(boolean show) {

    }

    @Override
    public boolean isDisplayRowColHeadings() {
        return false;
    }

    @Override
    public void setRowBreak(int row) {

    }

    @Override
    public boolean isRowBroken(int row) {
        return false;
    }

    @Override
    public void removeRowBreak(int row) {

    }

    @Override
    public int[] getRowBreaks() {
        return new int[0];
    }

    @Override
    public int[] getColumnBreaks() {
        return new int[0];
    }

    @Override
    public void setColumnBreak(int column) {

    }

    @Override
    public boolean isColumnBroken(int column) {
        return false;
    }

    @Override
    public void removeColumnBreak(int column) {

    }

    @Override
    public void setColumnGroupCollapsed(int columnNumber, boolean collapsed) {

    }

    @Override
    public void groupColumn(int fromColumn, int toColumn) {

    }

    @Override
    public void ungroupColumn(int fromColumn, int toColumn) {

    }

    @Override
    public void groupRow(int fromRow, int toRow) {

    }

    @Override
    public void ungroupRow(int fromRow, int toRow) {

    }

    @Override
    public void setRowGroupCollapsed(int row, boolean collapse) {

    }

    @Override
    public void setDefaultColumnStyle(int column, CellStyle style) {

    }

    @Override
    public void autoSizeColumn(int column) {

    }

    @Override
    public void autoSizeColumn(int column, boolean useMergedCells) {

    }

    @Override
    public Comment getCellComment(CellAddress ref) {
        return null;
    }

    @Override
    public void setActiveCell(CellAddress address) {}

    @Override
    public CellAddress getActiveCell() { return new CellAddress(0, 0); }

    @Override
    public java.util.List<? extends Hyperlink> getHyperlinkList() { return java.util.Collections.emptyList(); }

    @Override
    public Hyperlink getHyperlink(CellAddress address) { return null; }

    @Override
    public Hyperlink getHyperlink(int row, int column) { return null; }

    @Override
    public int getColumnOutlineLevel(int columnIndex) { return 0; }

    @Override
    public java.util.List<? extends DataValidation> getDataValidations() { return java.util.Collections.emptyList(); }

    @Override
    public java.util.Map<CellAddress, ? extends Comment> getCellComments() { return java.util.Collections.emptyMap(); }

    @Override
    public boolean isPrintRowAndColumnHeadings() { return false; }

    @Override
    public void setPrintRowAndColumnHeadings(boolean show) {}

    @Override
    public PaneInformation getPaneInformation() { return null; }

    @Override
    public Drawing<?> createDrawingPatriarch() {
        return null;
    }

    @Override
    public Workbook getWorkbook() {
        return null;
    }

    @Override
    public String getSheetName() {
        return "CSV";
    }

    @Override
    public boolean isSelected() {
        return false;
    }

    @Override
    public CellRange<? extends Cell> setArrayFormula(String formula, CellRangeAddress range) {
        return null;
    }

    @Override
    public CellRange<? extends Cell> removeArrayFormula(Cell cell) {
        return null;
    }

    @Override
    public DataValidationHelper getDataValidationHelper() {
        return null;
    }

    @Override
    public void addValidationData(DataValidation dataValidation) {

    }

    @Override
    public AutoFilter setAutoFilter(CellRangeAddress range) {
        return null;
    }

    @Override
    public SheetConditionalFormatting getSheetConditionalFormatting() {
        return null;
    }

    @Override
    public CellRangeAddress getRepeatingRows() {
        return null;
    }

    @Override
    public CellRangeAddress getRepeatingColumns() {
        return null;
    }

    @Override
    public void setRepeatingRows(CellRangeAddress rowRangeRef) {

    }

    @Override
    public void setRepeatingColumns(CellRangeAddress columnRangeRef) {

    }

    @Override
    public Iterator<Row> iterator() {
        return rowIterator();
    }

    @Override
    public Drawing<?> getDrawingPatriarch() { return null; }

    @Override
    public float getColumnWidthInPixels(int columnIndex) { return 0; }

    @Override
    public int addMergedRegionUnsafe(CellRangeAddress region) { return 0; }

    @Override
    public java.util.List<CellRangeAddress> getMergedRegions() { return java.util.Collections.emptyList(); }

    @Override
    public void removeMergedRegions(java.util.Collection<java.lang.Integer> indices) {}

    @Override
    public void validateMergedRegions() {}

    @Override
    public void createSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, PaneType activePane) {}

    @Override
    public double getMargin(PageMargin margin) { return 0; }

    @Override
    public void setMargin(PageMargin margin, double size) {}

    @Override
    public void shiftColumns(int startColumn, int endColumn, int n) {}
}
