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

import org.apache.poi.ss.usermodel.Workbook;

/**
 * {@link WorkbookStrategy} implementation for the modern XLSX format (Excel 2007+, OOXML).
 *
 * <p>Creates a fully in-memory {@link BusinessXSSFWorkbook}, which extends
 * {@link org.apache.poi.xssf.usermodel.XSSFWorkbook} with additional picture-management
 * support. The entire workbook is kept in heap memory, so this strategy is best suited for
 * moderate-sized exports; use {@link BigDataWorkbookStrategy} for large data sets.</p>
 *
 * @author dh
 * @since 1.0
 */
public class XlsxWorkbookStrategy implements WorkbookStrategy {

    /**
     * Creates a new in-memory {@link BusinessXSSFWorkbook} for XLSX output.
     *
     * @return a new {@link BusinessXSSFWorkbook} instance
     */
    @Override
    public Workbook createWorkbook() {
        return new BusinessXSSFWorkbook();
    }

    /**
     * Returns {@code "xlsx"}, the file-extension identifier for this format.
     *
     * @return {@code "xlsx"}
     */
    @Override
    public String getExcelType() {
        return "xlsx";
    }

    /**
     * Returns {@code false} because {@link BusinessXSSFWorkbook} is fully in-memory and
     * does not use streaming.
     *
     * @return {@code false}
     */
    @Override
    public boolean isBigData() {
        return false;
    }
}
