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

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;

/**
 * Builder for cascade dropdown models, supporting arbitrary-depth column cascade configuration.
 *
 * <h3>Usage examples</h3>
 * <pre>
 * // Single-column simple dropdown
 * CascadeValidateModelBuilder.builder("alarmLevel")
 *     .addItems(alarmLevelEnums, EnumDTO::getName)
 *     .build();
 *
 * // Two-level cascade (major category → minor category)
 * CascadeValidateModelBuilder.builder("devBigType")
 *     .addItem("AnalogDevice",
 *         CascadeValidateModelBuilder.builder("devSmallType")
 *             .addItems(smallTypes, EnumDTO::getName))
 *     .build();
 * </pre>
 *
 * @author dhsolo
 * @since 1.0
 */
public class CascadeValidateModelBuilder {

    private final CascadeValidateModel cascadeValidateModel;
    private List<CascadeValidateModel.CascadeValidateItem> items;

    /**
     * Creates a builder for the cascade column bound to the given Java field name.
     *
     * @param fieldName the field name of the column this dropdown is attached to
     */
    public CascadeValidateModelBuilder(String fieldName) {
        cascadeValidateModel = new CascadeValidateModel();
        cascadeValidateModel.setFieldName(fieldName);
    }

    /** Adds a single option with no child cascade. */
    public CascadeValidateModelBuilder addItem(String value) {
        return addItem(value, (CascadeValidateModelBuilder[]) null);
    }

    /**
     * Adds a single option, optionally with one or more child cascade builders.
     * <p><b>null-safe</b>: any {@code null} builder arguments are filtered out automatically,
     * allowing callers to pass optional child cascades via a ternary expression without extra null checks:
     * <pre>
     * builder.addItem(name, hasChildren ? childBuilder : null);
     * </pre>
     */
    public CascadeValidateModelBuilder addItem(String value, CascadeValidateModelBuilder... builders) {
        if (items == null) {
            items = new ArrayList<>();
        }
        var item = new CascadeValidateModel.CascadeValidateItem();
        item.setValue(value);
        if (builders != null && builders.length > 0) {
            var children = Arrays.stream(builders)
                    .filter(Objects::nonNull)
                    .map(CascadeValidateModelBuilder::build)
                    .toList();
            if (!children.isEmpty()) {
                item.setCascadeValidateModels(children);
            }
        }
        items.add(item);
        return this;
    }

    /**
     * Adds options in bulk from a collection (no child cascade), extracting display text via {@code nameMapper}.
     * <p>Replaces repetitive for-each loops:
     * <pre>
     * // Old style
     * for (EnumDTO e : enums) builder.addItem(e.getName());
     *
     * // New style
     * builder.addItems(enums, EnumDTO::getName);
     * </pre>
     *
     * @param items      data collection; returns this directly if {@code null}
     * @param nameMapper function to extract display text
     */
    public <T> CascadeValidateModelBuilder addItems(Collection<T> items, Function<T, String> nameMapper) {
        if (items != null) {
            items.forEach(item -> addItem(nameMapper.apply(item)));
        }
        return this;
    }

    /** Returns whether any options have been added (used to determine whether a child builder is needed). */
    public boolean hasItems() {
        return items != null && !items.isEmpty();
    }

    /**
     * Configures whether unrecognised values trigger a translation exception during import.
     *
     * @param needAddTranslationException {@code true} to raise a validation error; {@code false} to pass through
     * @return this builder
     */
    public CascadeValidateModelBuilder needAddTranslationException(boolean needAddTranslationException) {
        cascadeValidateModel.setNeedAddTranslationException(needAddTranslationException);
        return this;
    }

    /**
     * Finalises the builder and returns the configured {@link CascadeValidateModel}.
     *
     * @return the completed cascade validate model
     */
    public CascadeValidateModel build() {
        cascadeValidateModel.setItems(items);
        return cascadeValidateModel;
    }

    /**
     * Factory method — preferred entry point for starting a new cascade column builder.
     *
     * @param fieldName the Java field name of the column this dropdown is attached to
     * @return a new {@code CascadeValidateModelBuilder}
     */
    public static CascadeValidateModelBuilder builder(String fieldName) {
        return new CascadeValidateModelBuilder(fieldName);
    }
}
