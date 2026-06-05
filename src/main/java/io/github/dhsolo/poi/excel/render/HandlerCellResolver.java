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

/**
 * Custom handler resolver: applies value transformation via the {@link ExcelTranslateHandler} Function.
 * Corresponds to {@code @ExcelTranslateMethod}-driven translation methods and has higher priority
 * than all built-in resolvers.
 *
 * @author dh
 * @since 1.0
 */
public class HandlerCellResolver implements CellValueResolver {

    /**
     * Returns order value {@code 100}, giving this resolver higher priority than
     * {@link TranslateCellResolver} (200) and {@link PlainCellResolver} (MAX), but lower
     * than {@link PictureCellResolver} (50).
     */
    @Override
    public int getOrder() {
        return 100;
    }

    /**
     * Returns {@code true} when the column model declares a custom handler method via
     * {@link ExcelTranslateHandler#needHandle()}.
     *
     * @param model column metadata to inspect
     * @return {@code true} if a handler function is configured
     */
    @Override
    public boolean supports(ExcelModel model) {
        return ((ExcelTranslateHandler) model).needHandle();
    }

    /**
     * Invokes the column's custom handler with an {@link io.github.dhsolo.poi.excel.model.ExcelRowData}
     * snapshot and writes the returned value to the cell. Falls back to the column-level or
     * sheet-level default value when the handler returns {@code null}.
     *
     * @param ctx the render context for the current cell
     * @return always {@code 0} — handlers do not consume additional columns
     */
    @Override
    public int resolve(CellResolveContext ctx) {
        var excelRowData = ctx.cellValueSetter().createExcelRowData(
                ctx.dataObj(), ctx.model().getFieldName(), ctx.colIndex(), ctx.rowIndex());
        var result = ((ExcelTranslateHandler) ctx.model()).handler(excelRowData);
        if (result == null) {
            var v = Reflect.hasText(ctx.model().getNoneCellDefaultValue())
                    ? ctx.model().getNoneCellDefaultValue() : ctx.noneCellDefaultValue();
            ctx.cellValueSetter().setCellValue(ctx.cell(), v);
        } else {
            ctx.cellValueSetter().setCellValue(ctx.cell(), result);
        }
        return 0;
    }
}
