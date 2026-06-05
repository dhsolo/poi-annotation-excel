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

import org.apache.poi.ss.usermodel.Workbook;

/**
 * Strategy interface for workbook creation, following the Strategy design pattern.
 *
 * <p>Each implementation encapsulates the construction logic for a specific Apache POI
 * workbook type (XLS, XLSX, or streaming SXSSF for large data sets). The export pipeline
 * depends only on this interface, allowing the concrete format to be swapped at runtime
 * without changing downstream code.</p>
 *
 * @author dhsolo
 * @since 1.0
 */
public interface WorkbookStrategy {

    /**
     * Creates and returns a new {@link Workbook} instance appropriate for this strategy's
     * target format.
     *
     * @return a freshly constructed {@link Workbook}; never {@code null}
     */
    Workbook createWorkbook();

    /**
     * Returns the file-extension identifier for the format produced by this strategy.
     *
     * @return a lowercase extension string such as {@code "xls"} or {@code "xlsx"}
     */
    String getExcelType();

    /**
     * Indicates whether this strategy uses streaming (big-data) mode to avoid
     * {@link OutOfMemoryError} when writing large spreadsheets.
     *
     * @return {@code true} if the created workbook writes rows in a streaming fashion;
     *         {@code false} for fully in-memory workbooks
     */
    boolean isBigData();
}
