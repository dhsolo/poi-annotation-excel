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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * {@link WorkbookStrategy} implementation for the legacy XLS format (Excel 97-2003, BIFF8).
 *
 * <p>Creates a fully in-memory {@link HSSFWorkbook}. Because the entire workbook is held
 * in heap memory, this strategy is unsuitable for very large data sets. It is limited to
 * 65,535 rows per sheet as per the XLS specification.</p>
 *
 * @author dh
 * @since 1.0
 */
public class XlsWorkbookStrategy implements WorkbookStrategy {

    /**
     * Creates a new in-memory {@link HSSFWorkbook} for XLS output.
     *
     * @return a new {@link HSSFWorkbook} instance
     */
    @Override
    public Workbook createWorkbook() {
        return new HSSFWorkbook();
    }

    /**
     * Returns {@code "xls"}, the file-extension identifier for this format.
     *
     * @return {@code "xls"}
     */
    @Override
    public String getExcelType() {
        return "xls";
    }

    /**
     * Returns {@code false} because {@link HSSFWorkbook} is fully in-memory and does not
     * use streaming.
     *
     * @return {@code false}
     */
    @Override
    public boolean isBigData() {
        return false;
    }
}
