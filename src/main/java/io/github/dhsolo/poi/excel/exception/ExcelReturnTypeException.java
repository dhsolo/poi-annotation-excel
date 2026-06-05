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
 * Thrown when a method or field declares a return type that is incompatible with the
 * expected Excel cell value type for the current operation.
 *
 * <p>Examples of situations that trigger this exception:
 * <ul>
 *   <li>A custom {@link io.github.dhsolo.poi.excel.model.RowDataMapper} returns a type that cannot be
 *       coerced to a POI cell value (e.g. a complex domain object instead of a primitive,
 *       {@link String}, or {@link java.util.Date}).</li>
 *   <li>A getter annotated for export returns a collection where the framework expects
 *       a scalar value.</li>
 * </ul>
 *
 * @author dh
 * @since 1.0
 */
public class ExcelReturnTypeException extends ExcelException {

    /**
     * Constructs an {@code ExcelReturnTypeException} with no detail message.
     */
    public ExcelReturnTypeException() {
        super();
    }

    /**
     * Constructs an {@code ExcelReturnTypeException} with the given detail message.
     *
     * @param message description of the type mismatch, ideally including the actual and expected types
     */
    public ExcelReturnTypeException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code ExcelReturnTypeException} with a detail message and root cause.
     *
     * @param message description of the type mismatch
     * @param cause   the underlying exception; may be {@code null}
     */
    public ExcelReturnTypeException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs an {@code ExcelReturnTypeException} wrapping another throwable.
     *
     * @param cause the underlying cause; may be {@code null}
     */
    public ExcelReturnTypeException(Throwable cause) {
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
    public ExcelReturnTypeException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
