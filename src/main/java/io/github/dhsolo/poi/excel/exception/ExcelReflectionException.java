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
 * Thrown when a Java reflection operation performed on a model bean fails during
 * Excel import or export processing.
 *
 * <p>The framework relies heavily on reflection to read and write model fields at
 * runtime. This exception wraps lower-level reflection errors such as
 * {@link IllegalAccessException}, {@link java.lang.reflect.InvocationTargetException},
 * and {@link InstantiationException} so that callers receive a consistent, unchecked
 * exception from the Excel layer regardless of the underlying reflection API failure.
 *
 * @author dhsolo
 * @since 1.0
 */
public class ExcelReflectionException extends ExcelException {

    /**
     * Constructs an {@code ExcelReflectionException} with no detail message.
     */
    public ExcelReflectionException() {
        super();
    }

    /**
     * Constructs an {@code ExcelReflectionException} with the given detail message.
     *
     * @param message description of the reflection failure, ideally including the class
     *                and field or method name that could not be accessed
     */
    public ExcelReflectionException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code ExcelReflectionException} with a detail message and root cause.
     *
     * @param message description of the reflection failure
     * @param cause   the original reflection exception (e.g. {@link IllegalAccessException});
     *                may be {@code null}
     */
    public ExcelReflectionException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs an {@code ExcelReflectionException} wrapping another throwable.
     *
     * @param cause the underlying reflection exception; may be {@code null}
     */
    public ExcelReflectionException(Throwable cause) {
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
    protected ExcelReflectionException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
