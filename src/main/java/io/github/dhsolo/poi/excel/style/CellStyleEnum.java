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

import io.github.dhsolo.poi.excel.model.DiyRowContextCellModel;

/**
 * Identifies the logical cell-style role to apply when writing a cell during Excel export.
 *
 * <p>The framework maintains a small palette of pre-built {@link org.apache.poi.ss.usermodel.CellStyle}
 * objects per workbook. Rather than passing raw style objects throughout the rendering pipeline,
 * components reference one of these symbolic roles. The actual POI {@code CellStyle} instance is
 * resolved at write time from the current workbook's style registry.
 *
 * <p>When a cell should use a fully custom style not covered by the predefined roles, set the role
 * to {@link #customStyle} and supply the {@link org.apache.poi.ss.usermodel.CellStyle} directly via
 * {@link DiyRowContextCellModel#setCustomCellStyle(org.apache.poi.ss.usermodel.CellStyle)}.
 *
 * @author dh
 * @since 1.0
 */
public enum CellStyleEnum {

    /**
     * Style applied to title or header row cells — typically bold text, a coloured background,
     * and centred alignment to visually distinguish headers from data rows.
     */
    titleStyle,

    /**
     * Style applied to regular data cells in the body of the table — typically a standard
     * font with border lines and default alignment.
     */
    normalStyle,

    /**
     * Fallback style used when no specific role has been set.  The framework applies this
     * style when the cell has not been explicitly assigned {@code titleStyle}, {@code normalStyle},
     * or {@code customStyle}.
     */
    defaultStyle,

    /**
     * Marker indicating that the cell should use a caller-supplied
     * {@link org.apache.poi.ss.usermodel.CellStyle} instance rather than any pre-built style.
     * When this value is selected the companion custom-style field must be populated.
     */
    customStyle
}
