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
 * Translation resolver: translates the raw value through a mapping table before writing it to the cell.
 * <p>Value extraction is first attempted via {@link CellValueResolver#resolveSourceValue},
 * supporting {@code @ExcelColumn(sourcePath)} path traversal and {@code sourceField} field redirection.
 *
 * @author dh
 * @since 1.0
 */
public class TranslateCellResolver implements CellValueResolver {

    /**
     * Returns order value {@code 200}, lower priority than {@link HandlerCellResolver} (100)
     * and {@link PictureCellResolver} (50) but higher than {@link PlainCellResolver} (MAX).
     */
    @Override
    public int getOrder() {
        return 200;
    }

    /**
     * Returns {@code true} for non-picture columns that have a translation mapping table
     * configured (e.g., enum code → display label).
     *
     * @param model column metadata to inspect
     * @return {@code true} when translation is required and the column is not a picture
     */
    @Override
    public boolean supports(ExcelModel model) {
        return model.isNeedtranslate() && !model.isPicture();
    }

    /**
     * Extracts the raw value, looks it up in the column's translation map, and writes the
     * translated label to the cell. If the raw value is not found in the map it is written
     * as-is. Falls back to the column-level or sheet-level default when the value is
     * {@code null} after translation.
     *
     * @param ctx the render context for the current cell
     * @return always {@code 0} — translation does not consume extra columns
     */
    @Override
    public int resolve(CellResolveContext ctx) {
        var value = resolveSourceValue(ctx);
        var translate = ctx.model().getTranslateMappingInfo();
        if (translate != null && !translate.isEmpty() && value != null) {
            if (translate.containsKey(value)) {
                // Fast path: the raw value matches a key directly (equals/hashCode).
                value = translate.get(value);
            } else {
                String vs = value.toString();
                if (translate.containsKey(vs)) {
                    // Fast path: the value's string form matches a String key (the common case,
                    // e.g. Integer 0 against key "0").
                    value = translate.get(vs);
                } else {
                    // Fallback: rare maps with non-String keys matched by string form.
                    for (var key : translate.keySet()) {
                        if (key != null && key.toString().equals(vs)) {
                            value = translate.get(key);
                            break;
                        }
                    }
                }
            }
        }
        if (value == null) {
            value = Reflect.hasText(ctx.model().getNoneCellDefaultValue())
                    ? ctx.model().getNoneCellDefaultValue() : ctx.noneCellDefaultValue();
        }
        ctx.cellValueSetter().setCellValue(ctx.cell(), value,
                ctx.model().isDate() ? ctx.model().getPattern() : null);
        return 0;
    }
}
