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
package io.github.dh.poi.excel.cascade;

import java.util.List;

/**
 * Model describing one level of a cascade dropdown validation column.
 *
 * <p>A {@code CascadeValidateModel} represents a single column whose allowed values are
 * constrained by a dropdown list. Each option ({@link CascadeValidateItem}) may itself
 * carry child {@code CascadeValidateModel} instances, forming an arbitrarily deep tree
 * that drives multi-level dependent-dropdown validation in generated Excel files.
 *
 * <p>Instances are typically constructed via {@link CascadeValidateModelBuilder}.
 *
 * @author dh
 * @since 1.0
 */
public class CascadeValidateModel {

    /**
     * The dropdown column field name.
     */
    private String fieldName;

    /**
     * Whether a translation-lookup exception should be added for this column when the
     * selected value cannot be found in the translation map during import. Defaults to
     * {@code true} so that unrecognised values surface as validation errors rather than
     * being silently ignored.
     */
    private boolean isNeedAddTranslationException = true;

    /**
     * Dropdown options.
     */
    private List<CascadeValidateItem>  items;


    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }


    public List<CascadeValidateItem> getItems() {
        return items;
    }

    public void setItems(List<CascadeValidateItem> items) {
        this.items = items;
    }

    /**
     * One selectable option within a {@link CascadeValidateModel} dropdown column.
     *
     * <p>Each item holds the display/selection value and an optional list of child
     * {@link CascadeValidateModel} instances that become active when this item is chosen,
     * enabling multi-level cascade validation.
     */
    public static class CascadeValidateItem{

        /**
         * Dropdown value.
         */
        private String value;


        /**
         * Cascade dropdown column options.
         */
        private List<CascadeValidateModel> cascadeValidateModels;



        public String getValue() {
            return value;
        }

        public void setValue(String value) {
            this.value = value;
        }

        /**
         * Returns the child cascade models that are activated when this item is selected,
         * or {@code null} / an empty list if this is a leaf option.
         *
         * @return child cascade column models, may be {@code null}
         */
        public List<CascadeValidateModel> getCascadeValidateModels() {
            return cascadeValidateModels;
        }

        public void setCascadeValidateModels(List<CascadeValidateModel> cascadeValidateModels) {
            this.cascadeValidateModels = cascadeValidateModels;
        }
    }

    /**
     * Controls whether an exception is raised during import when the cell value cannot be
     * mapped through the translation table for this column.
     *
     * @param needAddTranslationException {@code true} to raise a validation error on
     *                                    unrecognised values; {@code false} to pass them through
     */
    public void setNeedAddTranslationException(boolean needAddTranslationException) {
        isNeedAddTranslationException = needAddTranslationException;
    }

    public boolean isNeedAddTranslationException() {
        return isNeedAddTranslationException;
    }
}
