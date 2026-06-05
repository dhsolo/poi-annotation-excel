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
package io.github.dhsolo.poi.excel.cascade;

import io.github.dhsolo.poi.excel.ExcelModel;
import java.util.List;

/**
 * Runtime wrapper around a {@link CascadeValidateModel} used while building the Excel sheet.
 *
 * <p>During export, the framework flattens the logical cascade tree into a list of
 * {@code CascadeValidateModelWrapper} nodes. Each node holds the resolved
 * {@link ExcelModel} for its column, the fully-expanded list of selectable
 * {@link CascadeValidateItemWrapper} options, an optional reference to the parent item
 * that activates this node, and flags that control validation behaviour.
 *
 * @author dh
 * @since 1.0
 */
public  class CascadeValidateModelWrapper {

    /**
     * Whether this column has dependent child cascade columns. When {@code true}, the
     * items of this column each carry their own child {@link CascadeValidateModelWrapper}
     * nodes which must also be written to the hidden data sheet.
     */
    private boolean hasCascade;

    /** The column metadata (header label, field binding, width, etc.) for this cascade column. */
    private ExcelModel excelModel;

    /**
     * Whether a translation-lookup exception should be raised when the imported cell value
     * cannot be mapped. Mirrors the flag on {@link CascadeValidateModel}.
     */
    private boolean isNeedAddTranslationException = true;

    /** The flattened, prefix-expanded list of selectable options for this column. */
    private List<CascadeValidateItemWrapper> items;

    /**
     * The parent item that gates the visibility of this column, or {@code null} for
     * top-level (unconditional) cascade columns.
     */
    private CascadeValidateItemWrapper parentValue;


    public boolean isHasCascade() {
        return hasCascade;
    }

    public void setHasCascade(boolean hasCascade) {
        this.hasCascade = hasCascade;
    }

    public List<CascadeValidateItemWrapper> getItems() {
        return items;
    }

    public void setItems(List<CascadeValidateItemWrapper> items) {
        this.items = items;
    }

    /**
     * Returns the parent {@link CascadeValidateItemWrapper} whose selection activates this
     * column, or {@code null} if this is a root-level cascade column.
     *
     * @return parent item wrapper, may be {@code null}
     */
    public CascadeValidateItemWrapper getParentValue() {
        return parentValue;
    }

    public void setParentValue(CascadeValidateItemWrapper parentValue) {
        this.parentValue = parentValue;
    }

    public ExcelModel getExcelModel() {
        return excelModel;
    }

    public void setExcelModel(ExcelModel excelModel) {
        this.excelModel = excelModel;
    }

    public void setNeedAddTranslationException(boolean needAddTranslationException) {
        isNeedAddTranslationException = needAddTranslationException;
    }

    public boolean isNeedAddTranslationException() {
        return isNeedAddTranslationException;
    }
}
