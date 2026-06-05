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

/**
 * Strategy interface for cell value resolution.
 * Strategy pattern: decouples different cell value population logic into independent strategy implementations.
 *
 * @author dh
 * @since 1.0
 */
public interface CellValueResolver {

    /**
     * Highest priority (smallest value, executed first).
     */
    int HIGHEST_PRECEDENCE = Integer.MIN_VALUE;

    /**
     * Lowest priority (largest value, executed last).
     */
    int LOWEST_PRECEDENCE = Integer.MAX_VALUE;

    /**
     * Returns whether this resolver can handle the given ExcelModel.
     */
    boolean supports(ExcelModel model);

    /**
     * Resolves and populates the cell value.
     *
     * @param context the render context containing the row, cell, data object, etc.
     * @return additional columns consumed (picture columns may span multiple columns); defaults to 0
     */
    int resolve(CellResolveContext context);

    /**
     * Returns the order priority of this resolver. Lower values have higher priority.
     * The resolver chain is sorted by this value to determine evaluation order.
     *
     * @return the order value (lower = higher priority)
     */
    default int getOrder() {
        return LOWEST_PRECEDENCE;
    }

    /**
     * Unified value extraction entry point: retrieves the raw value from the row data object
     * in priority order {@code sourcePath} → {@code sourceField} → {@code fieldName},
     * shared by all resolvers to avoid duplicated logic.
     */
    default Object resolveSourceValue(CellResolveContext ctx) {
        var model = ctx.model();
        if (model.hasSourcePath()) {
            return ctx.valueExtractor().getValueByPath(model.getSourcePath(), ctx.dataObj());
        }
        return ctx.valueExtractor().getValue(model.effectiveSourceField(), ctx.dataObj());
    }
}
