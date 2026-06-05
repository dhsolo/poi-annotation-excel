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

import java.util.List;

/**
 * Runtime wrapper around a single dropdown option within a cascade validation column.
 *
 * <p>During export, each raw {@link CascadeValidateModel.CascadeValidateItem} is
 * expanded into a {@code CascadeValidateItemWrapper} that carries:
 * <ul>
 *   <li>the display value written to the hidden data sheet;</li>
 *   <li>a back-reference to the {@link CascadeValidateModelWrapper} that owns this item;</li>
 *   <li>an {@code appendPrefix} flag indicating whether the item value was prefixed with
 *       its parent value to ensure uniqueness across cascade levels;</li>
 *   <li>the child {@link CascadeValidateModelWrapper} nodes activated when this item is selected.</li>
 * </ul>
 *
 * @author dhsolo
 * @since 1.0
 */
public class CascadeValidateItemWrapper {

    /** The cascade column model that contains this item. */
    private CascadeValidateModelWrapper ownModel;

    /**
     * The option value written to the hidden data sheet and used as the dropdown label.
     * May include a parent-value prefix when {@link #isAppendPrefix} is {@code true}.
     */
    private String value;

    /**
     * Whether the parent value was prepended to {@link #value} to guarantee uniqueness
     * across sibling groups. Excel's {@code INDIRECT} formula for cascade validation
     * requires each named range to be unique, so prefixing is applied when two sibling
     * groups share the same option text.
     */
    private boolean isAppendPrefix;

    /** Child cascade column wrappers that become active when this item is selected. */
    private List<CascadeValidateModelWrapper> cascadeValidateModels;

    /**
     * Returns the parent {@link CascadeValidateModelWrapper} that owns this item.
     *
     * @return owning model wrapper
     */
    public CascadeValidateModelWrapper getOwnModel() { return ownModel; }

    public String getValue() { return value; }

    /**
     * Returns the child cascade column wrappers activated by selecting this item,
     * or {@code null} / an empty list if this is a leaf option.
     *
     * @return child cascade model wrappers, may be {@code null}
     */
    public List<CascadeValidateModelWrapper> getCascadeValidateModels() { return cascadeValidateModels; }


    public void setOwnModel(CascadeValidateModelWrapper ownModel) {
        this.ownModel = ownModel;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public boolean isAppendPrefix() {
        return isAppendPrefix;
    }

    public void setAppendPrefix(boolean appendPrefix) {
        isAppendPrefix = appendPrefix;
    }

    public void setCascadeValidateModels( List<CascadeValidateModelWrapper> cascadeValidateModels) {
        this.cascadeValidateModels = cascadeValidateModels;
    }
}
