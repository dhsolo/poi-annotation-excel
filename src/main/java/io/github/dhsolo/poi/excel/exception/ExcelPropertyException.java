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
package io.github.dhsolo.poi.excel.exception;

/**
 * Thrown when a required bean property cannot be found, read, or written during
 * Excel import or export processing.
 *
 * <p>This exception typically surfaces when the framework attempts to read a field
 * value via reflection and the property name derived from an annotation does not
 * match any declared field or getter on the target class, or when the field exists
 * but is inaccessible (e.g. a private field in a sealed class hierarchy).
 *
 * @author dh
 * @since 1.0
 */
public class ExcelPropertyException extends ExcelException {

    /**
     * Constructs an {@code ExcelPropertyException} with no detail message.
     */
    public ExcelPropertyException() {
        super();
    }

    /**
     * Constructs an {@code ExcelPropertyException} with the given detail message.
     *
     * @param message description of which property could not be accessed and why
     */
    public ExcelPropertyException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code ExcelPropertyException} with a detail message and root cause.
     *
     * @param message description of the property access failure
     * @param cause   the underlying exception (e.g. {@link NoSuchFieldException}); may be {@code null}
     */
    public ExcelPropertyException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs an {@code ExcelPropertyException} wrapping another throwable.
     *
     * @param cause the underlying cause; may be {@code null}
     */
    public ExcelPropertyException(Throwable cause) {
        super(cause);
    }

    /**
     * Full-control constructor for subclassing or serialisation frameworks.
     *
     * @param message            the detail message
     * @param cause              the underlying cause; may be {@code null}
     * @param enableSuppression  whether suppressed exceptions can be added
     * @param writableStackTrace whether the stack trace should be filled in
     */
    public ExcelPropertyException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
