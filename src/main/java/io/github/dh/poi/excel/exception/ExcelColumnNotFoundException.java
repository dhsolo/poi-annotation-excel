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
package io.github.dh.poi.excel.exception;

/**
 * Thrown when a column referenced in an annotation or configuration cannot be located
 * in the actual Excel sheet being processed.
 *
 * <p>During import, the framework maps header cell text to model fields. This exception
 * is raised when a required column heading is absent from the spreadsheet, making it
 * impossible to populate the corresponding model field. During export, the same exception
 * may be raised if the column definition references a non-existent column index or name.
 *
 * @author dh
 * @since 1.0
 */
public class ExcelColumnNotFoundException extends ExcelException {

    /**
     * Constructs an {@code ExcelColumnNotFoundException} with no detail message.
     */
    public ExcelColumnNotFoundException() {
        super();
    }

    /**
     * Constructs an {@code ExcelColumnNotFoundException} with the given detail message.
     *
     * @param message description of which column could not be found, ideally including
     *                the column name or index and the sheet being processed
     */
    public ExcelColumnNotFoundException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code ExcelColumnNotFoundException} with a detail message and root cause.
     *
     * @param message description of the missing column
     * @param cause   the underlying exception; may be {@code null}
     */
    public ExcelColumnNotFoundException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs an {@code ExcelColumnNotFoundException} wrapping another throwable.
     *
     * @param cause the underlying cause; may be {@code null}
     */
    public ExcelColumnNotFoundException(Throwable cause) {
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
    protected ExcelColumnNotFoundException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
