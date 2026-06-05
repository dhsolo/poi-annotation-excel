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

import io.github.dhsolo.poi.excel.cascade.CascadeValidateModel;
import java.util.List;

/**
 * Marker-and-provider interface for Excel export model classes that need
 * cascading (linked) dropdown validation in generated sheets.
 *
 * <p>When an export model implements this interface, the export pipeline inspects the
 * returned {@link CascadeValidateModel} list and renders multi-level dependent dropdowns
 * (e.g. province → city → district) as Excel data-validation constraints on the
 * appropriate columns.</p>
 *
 * <p>Implementing classes typically build the cascade definition from their own field
 * metadata or from an externally injected data source.</p>
 *
 * @author dhsolo
 * @since 1.0
 * @see CascadeValidateModel
 */
public interface ExcelCascadeAble {

    /**
     * Returns the ordered list of cascade (linked-dropdown) validation definitions for
     * the Excel sheet generated from this model.
     *
     * <p>Each {@link CascadeValidateModel} in the list describes one dropdown column,
     * including its allowed values and any nested child dropdowns that are dependent on
     * the selected value.</p>
     *
     * @return a non-{@code null} list of cascade validation models; may be empty if no
     *         cascade validation is required
     */
    List<CascadeValidateModel> cascadeList();
}
