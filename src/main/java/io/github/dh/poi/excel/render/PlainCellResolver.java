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
package io.github.dh.poi.excel.render;

import io.github.dh.poi.excel.ExcelModel;
import io.github.dh.common.Reflect;

/**
 * Plain value resolver: reads a value directly from the object and writes it to the cell,
 * with no translation, no picture handling, and no custom handler.
 * <p>Supports {@code @ExcelColumn(sourcePath)} and {@code @ExcelColumn(sourceField)} path-based
 * value extraction via the inherited default {@link CellValueResolver#resolveSourceValue} implementation.
 *
 * @author dh
 * @since 1.0
 */
public class PlainCellResolver implements CellValueResolver {

    /**
     * Returns {@link CellValueResolver#LOWEST_PRECEDENCE}, acting as the final fallback
     * after all higher-priority resolvers have declined.
     */
    @Override
    public int getOrder() {
        return LOWEST_PRECEDENCE;
    }

    /**
     * Returns {@code true} when the column requires neither translation nor a custom handler
     * and is not a picture column — i.e., it is a straightforward field-to-cell mapping.
     *
     * @param model column metadata to inspect
     * @return {@code true} for plain value columns
     */
    @Override
    public boolean supports(ExcelModel model) {
        return !model.isNeedtranslate()
                && !((ExcelTranslateHandler) model).needHandle()
                && !model.isPicture();
    }

    /**
     * Extracts the field value from the row data object via
     * {@link CellValueResolver#resolveSourceValue} and writes it directly to the cell.
     * Falls back to the column-level or sheet-level default when the value is {@code null}.
     *
     * @param ctx the render context for the current cell
     * @return always {@code 0} — plain cells never consume extra columns
     */
    @Override
    public int resolve(CellResolveContext ctx) {
        var value = resolveSourceValue(ctx);
        if (value == null) {
            var defaultVal = Reflect.hasText(ctx.model().getNoneCellDefaultValue())
                    ? ctx.model().getNoneCellDefaultValue() : ctx.noneCellDefaultValue();
            ctx.cellValueSetter().setCellValue(ctx.cell(), defaultVal);
        } else {
            ctx.cellValueSetter().setCellValue(ctx.cell(), value,
                    ctx.model().isDate() ? ctx.model().getPattern() : null);
        }
        return 0;
    }
}
