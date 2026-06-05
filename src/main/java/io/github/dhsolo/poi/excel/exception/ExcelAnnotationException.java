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
 * Thrown when a required Excel annotation is missing, misconfigured, or applied to an
 * incompatible element.
 *
 * <p>Typical triggers include:
 * <ul>
 *   <li>A model class is missing the {@code @ExcelInfo} annotation that the framework requires.</li>
 *   <li>An {@code @ExcelColumn} annotation references a column that cannot be resolved.</li>
 *   <li>Conflicting annotation attributes are detected during pre-processing.</li>
 * </ul>
 *
 * @author dhsolo
 * @since 1.0
 */
public class ExcelAnnotationException extends ExcelException {

    /**
     * Constructs an {@code ExcelAnnotationException} with no detail message.
     */
    public ExcelAnnotationException() {
        super();
    }

    /**
     * Constructs an {@code ExcelAnnotationException} with the given detail message.
     *
     * @param message explanation of which annotation is missing or invalid
     */
    public ExcelAnnotationException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code ExcelAnnotationException} with a detail message and root cause.
     *
     * @param message explanation of the annotation problem
     * @param cause   the underlying exception that triggered this error; may be {@code null}
     */
    public ExcelAnnotationException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs an {@code ExcelAnnotationException} wrapping another throwable.
     *
     * @param cause the underlying cause; may be {@code null}
     */
    public ExcelAnnotationException(Throwable cause) {
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
    public ExcelAnnotationException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
