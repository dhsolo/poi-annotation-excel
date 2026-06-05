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

import io.github.dhsolo.poi.excel.ExcelModel;
import io.github.dhsolo.common.Reflect;
import io.github.dhsolo.poi.excel.picture.PictureHandler;

/**
 * Picture resolver: handles downloading picture URLs, embedding them, and expanding multi-picture columns.
 *
 * @author dh
 * @since 1.0
 */
public class PictureCellResolver implements CellValueResolver {

    /**
     * Returns order value {@code 50}, the highest priority among the built-in resolvers,
     * ensuring picture columns are handled before translation or plain-value resolvers.
     */
    @Override
    public int getOrder() {
        return 50;
    }

    /**
     * Returns {@code true} when the column is declared as a picture column.
     *
     * @param model column metadata to inspect
     * @return {@code true} if this is a picture column
     */
    @Override
    public boolean supports(ExcelModel model) {
        return model.isPicture();
    }

    /**
     * Downloads the image identified by the cell's URL value, embeds it at the cell
     * position, and creates additional blank styled cells for any extra columns required
     * when multiple images share the same column.
     *
     * <p>When the value is {@code null} the cell receives the configured default text and
     * the method still delegates to {@link PictureHandler#setPicture} so that row/column
     * dimension tracking remains consistent.
     *
     * @param ctx the render context for the current cell
     * @return the number of extra columns consumed by picture expansion (0 when no
     *         multi-image expansion occurs)
     */
    @Override
    public int resolve(CellResolveContext ctx) {
        var value = ctx.valueExtractor().getValue(ctx.model().getFieldName(), ctx.dataObj());
        if (value == null) {
            var v = Reflect.hasText(ctx.model().getNoneCellDefaultValue())
                    ? ctx.model().getNoneCellDefaultValue() : ctx.noneCellDefaultValue();
            ctx.cellValueSetter().setCellValue(ctx.cell(), v);
        }

        var imageUrl = value != null ? value.toString() : null;
        int max = ctx.pictureHandler().setPicture(
                ctx.colIndex(), ctx.rowNum() + ctx.rowIndex(),
                ctx.colIndex(), ctx.rowNum() + ctx.rowIndex(),
                imageUrl, ctx.cell());

        var columnMaxNum = ctx.pictureHandler().getColumnMaxMapping().get(ctx.colIndex());
        if (columnMaxNum != null) {
            max = columnMaxNum;
            int addIndex = max;
            int add = ctx.colIndex() + 1;
            while (addIndex-- > 0) {
                var extraCell = ctx.row().createCell(add);
                extraCell.setCellStyle(ctx.dataCellStyle());
                add++;
            }
        }
        return max;
    }
}
