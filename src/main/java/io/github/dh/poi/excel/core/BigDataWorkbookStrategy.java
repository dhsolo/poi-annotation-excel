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
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * {@link WorkbookStrategy} implementation for large-data XLSX exports using the SXSSF
 * streaming API.
 *
 * <p>Creates a {@link BusinessSXSSFWorkbook} backed by a {@link BusinessXSSFWorkbook}.
 * Unlike the in-memory {@link XlsxWorkbookStrategy}, this strategy flushes rows to a
 * temporary disk file as they are written, keeping only a configurable sliding window of
 * rows in heap memory. This prevents {@link OutOfMemoryError} when exporting hundreds of
 * thousands of rows. The output format is still XLSX.</p>
 *
 * <p>Note: random access to already-flushed rows is not supported by the SXSSF API.</p>
 *
 * @author dh
 * @since 1.0
 * @see BusinessSXSSFWorkbook
 * @see SXSSFWorkbook
 */
public class BigDataWorkbookStrategy implements WorkbookStrategy {

    /**
     * Creates a new streaming {@link BusinessSXSSFWorkbook} wrapping a
     * {@link BusinessXSSFWorkbook} for large-data XLSX output.
     *
     * @return a new {@link BusinessSXSSFWorkbook} instance
     */
    @Override
    public Workbook createWorkbook() {
        return new BusinessSXSSFWorkbook(new BusinessXSSFWorkbook());
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
     * Returns {@code true} to indicate that this strategy uses streaming (big-data) mode.
     *
     * @return {@code true}
     */
    @Override
    public boolean isBigData() {
        return true;
    }
}
