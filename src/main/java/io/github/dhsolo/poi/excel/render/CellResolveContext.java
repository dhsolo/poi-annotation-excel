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
package io.github.dhsolo.poi.excel.render;

import io.github.dhsolo.poi.excel.core.CellValueSetter;
import io.github.dhsolo.poi.excel.core.ValueExtractor;
import io.github.dhsolo.poi.excel.ExcelModel;
import io.github.dhsolo.poi.excel.picture.PictureHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Cell resolution context carrying all information needed for a single cell render pass.
 *
 * <p>Uses a JDK 17 record instead of the previous final class: the compiler automatically
 * generates the constructor, accessors, equals/hashCode/toString, eliminating boilerplate
 * while guaranteeing instance immutability.
 *
 * @param cellValueSetter      utility that writes a typed value into a POI {@link Cell}
 * @param valueExtractor       utility that reads a field value from the row data object
 * @param cell                 the target POI cell being populated
 * @param row                  the POI row that owns {@code cell}
 * @param sheet                the POI sheet being written
 * @param model                column metadata describing how this cell should be rendered
 * @param dataObj              the row data object from which the cell value is extracted
 * @param colIndex             0-based physical column index of {@code cell} in the sheet
 * @param rowIndex             0-based data-row index (offset from the first data row)
 * @param rowNum               absolute 0-based row number of the first data row in the sheet;
 *                             the cell's sheet row is {@code rowNum + rowIndex}
 * @param noneCellDefaultValue sheet-level fallback written when the resolved value is {@code null}
 *                             and the column has no column-level default of its own
 * @param dataCellStyle        shared {@link CellStyle} applied to ordinary data cells and to
 *                             any extra cells created by multi-picture expansion
 * @param pictureHandler       handler responsible for downloading and embedding images;
 *                             may be a no-op implementation when no picture columns are present
 * @author dhsolo
 * @since 1.0
 */
public record CellResolveContext(
        CellValueSetter cellValueSetter,
        ValueExtractor valueExtractor,
        Cell cell,
        Row row,
        Sheet sheet,
        ExcelModel model,
        Object dataObj,
        int colIndex,
        int rowIndex,
        int rowNum,
        String noneCellDefaultValue,
        CellStyle dataCellStyle,
        PictureHandler pictureHandler
) {}
