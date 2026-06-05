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
package io.github.dhsolo.poi.excel;

/**
 * Enumeration of the Excel file formats supported by this framework.
 *
 * <p>The format determines which Apache POI workbook implementation is used internally:
 * {@code XLSX} maps to {@code XSSFWorkbook} (Office Open XML, supports up to ~1 million rows),
 * while {@code XLS} maps to {@code HSSFWorkbook} (legacy BIFF8 format, limited to 65 535 rows).
 * Choose {@code XLSX} for new files unless compatibility with Excel 97–2003 is required.
 *
 * @author dh
 * @since 1.0
 */
public enum ExcelType {

    /**
     * Office Open XML spreadsheet format ({@code .xlsx}), introduced in Excel 2007.
     * Backed by {@code XSSFWorkbook}. Supports larger datasets and richer features
     * such as data validation lists and conditional formatting.
     */
    XLSX("xlsx"),

    /**
     * Legacy binary Excel format ({@code .xls}), compatible with Excel 97–2003.
     * Backed by {@code HSSFWorkbook}. Maximum of 65 535 rows and 256 columns per sheet.
     */
    XLS("xls");

    /** The lowercase file extension string (without leading dot) used to identify this format. */
    private final String value;

    ExcelType(String value) {
        this.value = value;
    }

    /**
     * Returns the lowercase file extension string for this format (e.g. {@code "xlsx"}).
     *
     * @return the file extension without a leading dot
     */
    public String getValue() {
        return this.value;
    }

    /**
     * Resolves an {@code ExcelType} from its file-extension string, ignoring case.
     *
     * @param value the extension string to look up, e.g. {@code "xlsx"} or {@code "XLS"}
     * @return the matching {@code ExcelType}
     * @throws IllegalArgumentException if {@code value} does not correspond to a known format
     */
    public static ExcelType fromValue(String value) {
        for (ExcelType type : values()) {
            if (type.value.equalsIgnoreCase(value)) {
                return type;
            }
        }
        throw new IllegalArgumentException("Unknown excel type: " + value);
    }
}
