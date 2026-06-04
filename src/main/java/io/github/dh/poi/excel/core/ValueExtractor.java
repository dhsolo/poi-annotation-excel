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
package io.github.dh.poi.excel.core;

import io.github.dh.poi.excel.exception.ExcelReflectionException;

/**
 * Strategy interface for reading a named property value from an arbitrary data object.
 *
 * <p>Follows the Spring {@code BeanWrapper} pattern: the picture and render packages depend
 * only on this narrow interface rather than on the full {@code ExcelCreator}, keeping the
 * dependency graph acyclic and the components independently testable.
 *
 * <p>The primary method {@link #getValue(String, Object)} resolves a single-level property
 * name (typically a Java field name or getter-derived name) from the given object.  The
 * default method {@link #getValueByPath(String, Object)} builds on top of this to support
 * dot-separated property paths for nested objects.
 *
 * @author dh
 * @since 1.0
 */
public interface ValueExtractor {

    /**
     * Extracts the value of the named property from the given data object.
     *
     * <p>Implementations typically use reflection or a pre-built method handle to call
     * the appropriate getter.  The property name convention follows the {@code @ExcelColumn}
     * field name used in the model class.
     *
     * @param field the property name to read (e.g. {@code "userName"})
     * @param data  the object from which to read the property; must not be {@code null}
     * @return the property value, or {@code null} if the property is {@code null}
     * @throws ExcelReflectionException if the property
     *         cannot be accessed via reflection
     */
    Object getValue(String field, Object data);

    /**
     * Extracts a value from a data object by following a dot-separated property path.
     *
     * <p>Each dot-delimited segment is resolved one level at a time by delegating to
     * {@link #getValue(String, Object)}.  If any intermediate object in the path is
     * {@code null}, the method short-circuits and returns {@code null} immediately rather
     * than throwing a {@code NullPointerException}.
     *
     * <p>Example: the path {@code "pointInfo.measureType"} is equivalent to
     * {@code data.getPointInfo().getMeasureType()}.
     *
     * @param path the dot-separated property path to traverse (e.g. {@code "address.city"})
     * @param data the root object from which to start traversal; must not be {@code null}
     * @return the value at the end of the path, or {@code null} if any level is {@code null}
     */
    default Object getValueByPath(String path, Object data) {
        var parts = path.split("\\.", -1);
        Object current = data;
        for (var part : parts) {
            if (current == null) return null;
            current = getValue(part, current);
        }
        return current;
    }
}
