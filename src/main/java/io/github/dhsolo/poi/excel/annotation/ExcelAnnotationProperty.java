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
package io.github.dhsolo.poi.excel.annotation;

import io.github.dhsolo.poi.excel.ExcelModel;
import io.github.dhsolo.poi.excel.cascade.CascadeValidateModel;

import java.util.List;
import java.util.Map;

/**
 * Read-only view of the annotation metadata extracted from a model class annotated with
 * {@code @ExcelInfo} and {@code @ExcelColumn}.
 *
 * <p>An instance of this interface is produced by the framework's annotation processor for
 * every model class and acts as the single source of truth for all sheet-level and
 * column-level configuration during both export and import.  Components in the rendering
 * and validation pipeline depend on this interface rather than on the raw annotations,
 * keeping them decoupled from reflection details.
 *
 * @author dhsolo
 * @since 1.0
 */
public interface ExcelAnnotationProperty {

    /**
     * Returns the sheet title text declared via {@code @ExcelInfo#title()}.
     * The title is written above the header row as a merged, prominently styled cell.
     *
     * @return the sheet title, or an empty string if no title was declared
     */
    String getTitle();

    /**
     * Returns the column header labels in display order, as declared by the
     * {@code @ExcelColumn#value()} attributes on the model class.
     *
     * @return an array of header strings; never {@code null}, but may be empty
     */
    String[] getHeader();

    /**
     * Returns the ordered list of column-level metadata models derived from the
     * {@code @ExcelColumn} annotations on the model class fields.
     *
     * @return a non-null, ordered list of {@link ExcelModel} descriptors
     */
    List<ExcelModel> getExcelModels();

    /**
     * Returns the class-level {@link ExcelInfo} annotation instance, giving access to
     * sheet-scoped settings such as the sheet name, freeze-pane configuration, and
     * whether an order-number column should be prepended.
     *
     * @return the {@code @ExcelInfo} annotation; never {@code null} for a valid model class
     */
    ExcelInfo getExcelInfo();

    /**
     * Returns the raw data object (typically a {@code List}) that will be iterated to
     * populate the sheet during export.
     *
     * @return the bound data object; may be {@code null} before data is attached
     */
    Object getExcelData();

    /**
     * Returns a map from zero-based column index to a merge-group identifier string.
     * Columns that share the same identifier value are eligible for vertical cell merging
     * when adjacent rows carry identical cell content.
     *
     * @return a non-null map; empty if no merge configuration was declared
     */
    Map<Integer, String> getMergeInfo();

    /**
     * Returns a map from zero-based column index to the desired column width in units
     * of 1/256 of a character width (the POI unit for column widths).
     *
     * @return a non-null map; empty if no explicit widths were declared
     */
    Map<Integer, Integer> getColumnWidthInfo();

    /**
     * Returns the cascade (linked) drop-down validation models declared on the model class.
     * Each entry describes a parent-child column pair whose valid values are dynamically
     * filtered based on the parent column's selected value.
     *
     * @return a non-null list of cascade validation descriptors; empty if none were declared
     */
    List<CascadeValidateModel> getCascadeValidateModel();

    /**
     * Returns the list of {@link AnnotationProcessor} instances generated for each
     * sub-model in a complex (multi-table) sheet.  For simple single-table sheets this
     * list is empty.
     *
     * @return a non-null list of complex sub-model processors; empty for simple sheets
     */
    List<AnnotationProcessor> getComplexExcelModels();

    /**
     * Returns the list of custom-row configurations declared by {@code @ExcelRow}
     * annotations, sorted in ascending order by {@code order}.  Each entry describes
     * a row that is inserted after the title row and before the column header row.
     *
     * @return a non-null, ordered list of {@link DiyRowConfig} entries; empty if no
     *         {@code @ExcelRow} annotations are present
     */
    List<DiyRowConfig> getDiyRows();

    /**
     * Immutable configuration for a single custom row derived from an {@code @ExcelRow} annotation.
     *
     * <p>Uses a JDK 17 record; the accessors {@link #text()} and {@link #merge()} are
     * generated automatically, eliminating boilerplate.
     *
     * @param text  the display text to write into the row's first cell
     * @param merge {@code true} if all cells in the row should be merged into a single
     *              cell spanning the full column width, matching the {@code needMerge}
     *              parameter of the underlying {@code addDiyRowContext} call
     */
    record DiyRowConfig(String text, boolean merge) {}

    /**
     * Returns the parent-header groups declared via {@code @ExcelColumnParent}, or {@code null}
     * when none are present. Column indices are in data-column space (excluding any order column).
     *
     * @return the parent-header spans, or {@code null}
     */
    default List<ParentHeader> getParentHeaders() { return null; }

    /**
     * A grouped (parent) header cell spanning a contiguous range of data columns.
     *
     * @param label    the parent header text
     * @param startCol the first data-column index covered (inclusive)
     * @param endCol   the last data-column index covered (inclusive)
     */
    record ParentHeader(String label, int startCol, int endCol) {}
}
