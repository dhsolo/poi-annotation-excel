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
package io.github.dhsolo.poi.excel.importor;

/**
 * Listener interface for row-by-row Excel import processing.
 *
 * <p>Implement this interface to receive parsed rows as they are read, avoiding the need
 * to hold all rows in memory. Similar in spirit to EasyExcel's {@code ReadListener}.
 *
 * <p>Example:
 * <pre>
 * ExcelImportor importer = new ExcelImportor(inputStream);
 * importer.addColumnName(columns);
 * importer.setReadListener(new ExcelReadListener&lt;MyDTO&gt;() {
 *     public void onRow(Map&lt;String, Object&gt; row, int rowIndex) {
 *         process(row);
 *     }
 *     public void onError(String message, int rowIndex) {
 *         log.warn("Row {} error: {}", rowIndex, message);
 *     }
 * });
 * importer.analysisExcel();
 * </pre>
 *
 * @author dh
 * @since 1.0
 */
public interface ExcelReadListener {

    /**
     * Called for each successfully parsed data row.
     *
     * @param row      field-name → cell-value map for this row
     * @param rowIndex 0-based index within the sheet (after the start row)
     */
    void onRow(java.util.Map<String, Object> row, int rowIndex);

    /**
     * Called when a validation or conversion error occurs on a row.
     * Default implementation does nothing (errors accumulate in {@code getErrorMessage()}).
     *
     * @param message  human-readable error description
     * @param rowIndex 0-based index within the sheet
     */
    default void onError(String message, int rowIndex) {}

    /**
     * Called once after all sheets are processed.
     *
     * @param totalRows total number of successfully parsed rows across all sheets
     */
    default void onFinish(int totalRows) {}
}
