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

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Strategy interface for managing Excel cell styles.
 *
 * <p>Maintains three canonical style roles — title, header, and data cell — and provides
 * copy methods that create independent style instances derived from each role. Using copies
 * rather than the shared instances prevents one caller's modifications from affecting others.
 *
 * <p>Implemented by {@link DefaultCellStyleManager}.
 *
 * @author dhsolo
 * @since 1.0
 */
public interface CellStyleManager {

    /**
     * Creates and registers the default title, header, and data cell styles on the
     * workbook. Must be called once before any style accessor or copy method is used.
     */
    void initDefaultStyles();

    /**
     * Returns the shared title cell style (largest font, bold).
     * Do not mutate the returned instance; use {@link #copyTitleStyle()} for per-cell customisation.
     *
     * @return shared title {@link CellStyle}
     */
    CellStyle getTitleCellStyle();

    /**
     * Returns the shared data cell style (standard body font, thin borders, centred).
     * Do not mutate the returned instance; use {@link #copyCellStyle()} for per-cell customisation.
     *
     * @return shared data {@link CellStyle}
     */
    CellStyle getCellStyle();

    /**
     * Returns the shared header cell style (medium font, bold, thin borders, centred).
     * Do not mutate the returned instance; use {@link #copyHeaderStyle()} for per-cell customisation.
     *
     * @return shared header {@link CellStyle}
     */
    CellStyle getHeaderCellStyle();

    /**
     * Replaces the shared title style with a custom one. Subsequent calls to
     * {@link #getTitleCellStyle()} and {@link #copyTitleStyle()} will use the new style.
     *
     * @param titleCellStyle replacement title style
     */
    void setTitleCellStyle(CellStyle titleCellStyle);

    /**
     * Replaces the shared data cell style with a custom one.
     *
     * @param cellStyle replacement data cell style
     */
    void setCellStyle(CellStyle cellStyle);

    /**
     * Replaces the shared header cell style with a custom one.
     *
     * @param headerCellStyle replacement header style
     */
    void setHeaderCellStyle(CellStyle headerCellStyle);

    /**
     * Creates a new independent {@link CellStyle} cloned from the current title style.
     * The caller may freely modify the returned instance.
     *
     * @return a modifiable copy of the title style
     */
    CellStyle copyTitleStyle();

    /**
     * Creates a new independent {@link CellStyle} cloned from the current data cell style.
     * The caller may freely modify the returned instance.
     *
     * @return a modifiable copy of the data cell style
     */
    CellStyle copyCellStyle();

    /**
     * Creates a new independent {@link CellStyle} cloned from the current header style.
     * The caller may freely modify the returned instance.
     *
     * @return a modifiable copy of the header style
     */
    CellStyle copyHeaderStyle();
}
