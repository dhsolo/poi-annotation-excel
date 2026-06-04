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
package io.github.dh.poi.excel.model;

import java.util.List;

/**
 * Marker interface for model classes whose single worksheet contains multiple independent
 * header-and-data sections stacked vertically.
 *
 * <p>In a standard export, every sheet has exactly one header row followed by one data set.
 * Implement {@code ComplexExcelModel} when a single sheet must display several distinct
 * tables one after another — for example, a summary table at the top followed by a
 * detail table further down.  The framework calls {@link #getComplexModels()} to retrieve
 * the ordered list of sub-models and renders each one sequentially, inserting headers and
 * data rows for every entry.
 *
 * <p>Each element returned by {@link #getComplexModels()} must itself be a valid Excel
 * model object (annotated with {@code @ExcelInfo} and {@code @ExcelColumn}) so that the
 * framework can derive column metadata via the standard annotation-processing path.
 *
 * @author dh
 * @since 1.0
 */
public interface ComplexExcelModel {

    /**
     * Returns the ordered list of sub-model objects to be rendered on the same sheet.
     *
     * <p>The framework iterates this list from first to last, writing each model's header
     * row and data rows in turn.  The list must not be {@code null}; an empty list results
     * in a blank sheet.
     *
     * @return an ordered, non-null list of sub-model objects; each element must carry the
     *         {@code @ExcelInfo} annotation so the framework can derive its column layout
     */
    List getComplexModels();
}
