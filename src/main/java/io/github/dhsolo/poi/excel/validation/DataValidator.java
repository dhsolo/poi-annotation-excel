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
package io.github.dhsolo.poi.excel.validation;

import io.github.dhsolo.poi.excel.ExcelModel;

import java.util.Map;

/**
 * Strategy interface for Excel data validation.
 *
 * <p>Handles dropdown lists, cascade validation, formula validation, etc.
 * Implemented by {@link DefaultDataValidator}.
 *
 * @author dhsolo
 * @since 1.0
 */
public interface DataValidator {

    /**
     * Sets how many data rows (counted from the first data row) dropdown validations and
     * formula pre-fill should cover. Implementations default to 1000 when unset.
     *
     * @param rowCount the number of data rows to cover; values &lt; 1 are ignored
     */
    default void setValidationRowCount(int rowCount) {
        // default: implementation keeps its built-in coverage
    }

    /**
     * Sets the index of the hidden data sheet used to store dropdown list source values.
     * Each validator instance is bound to exactly one list sheet; this method allows the
     * sheet index to be updated when multiple sheets are being generated.
     *
     * @param currentListNum 0-based index of the current list (hidden) sheet
     */
    void setCurrentListNum(int currentListNum);

    /**
     * Returns the 0-based index of the hidden data sheet that this validator writes
     * dropdown source values into.
     *
     * @return current list sheet index
     */
    int getCurrentListNum();

    /**
     * Sets the total number of data rows for which validation rules must be applied.
     * This controls how many rows are covered by each {@code DVConstraint}.
     *
     * @param rowNum total number of data rows in the target sheet
     */
    void setRowNum(int rowNum);

    /**
     * Applies all dropdown and cascade validation constraints to the target sheet.
     *
     * <p>Iterates over {@code columnMappingInfo}, writes option values to the hidden
     * list sheet, and registers {@code DataValidation} rules on each applicable column
     * for rows {@code 1} through {@code rowNum}.
     *
     * @param columnMappingInfo column-index-to-{@link ExcelModel} mapping for the sheet
     * @param needOrderNum      whether a leading order-number column is present (shifts
     *                          column indices by one when {@code true})
     * @param rowNum            number of data rows to which validation is applied
     */
    void checkListBox(Map<Integer, ExcelModel> columnMappingInfo, boolean needOrderNum, int rowNum);
}
