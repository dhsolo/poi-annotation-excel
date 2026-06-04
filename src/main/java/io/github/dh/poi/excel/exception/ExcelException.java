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
 * Base unchecked exception for all Excel-related operations in this framework.
 *
 * <p>Follows the Spring Framework {@code DataAccessException} pattern by providing
 * a common superclass for the entire Excel utility exception hierarchy. Every
 * framework-specific exception extends this class so callers can choose to catch
 * either this base type or a more specific subtype.
 *
 * <p>When constructed with an {@code errorCode}, the exception message is serialised
 * as a compact JSON string of the form {@code {"code":&lt;code&gt;,"message":"&lt;msg&gt;"}}
 * to make it convenient for HTTP error responses.
 *
 * @author dh
 * @since 1.0
 */
public class ExcelException extends RuntimeException {

    /** Numeric error code carried by this exception; defaults to 500 for message-only constructors. */
    private int code;

    /** Human-readable error description, mirroring {@link #getMessage()} for structured access. */
    private String message;

    /**
     * Constructs an {@code ExcelException} with no detail message.
     */
    public ExcelException() {
        super();
    }

    /**
     * Constructs an {@code ExcelException} with the given detail message.
     * The internal error code is set to {@code 500}.
     *
     * @param message the detail message
     */
    public ExcelException(String message) {
        super(message);
        this.code = 500;
    }

    /**
     * Constructs an {@code ExcelException} whose detail message is produced by
     * {@link String#format(String, Object[])} with the supplied format string and arguments.
     *
     * @param format a {@link java.util.Formatter} format string
     * @param args   arguments referenced by the format specifiers
     */
    public ExcelException(String format, Object... args) {
        super(String.format(format, args));
    }

    /**
     * Constructs an {@code ExcelException} carrying a structured error code and message.
     * The detail message is serialised as {@code {"code":<errorCode>,"message":"<message>"}}.
     *
     * @param errorCode numeric error code (e.g. an HTTP status or application error code)
     * @param message   human-readable error description
     */
    public ExcelException(int errorCode, String message) {
        this(String.format("{\"code\":%s,\"message\":\"%s\"}", errorCode, message));
        this.code = errorCode;
        this.message = message;
    }

    /**
     * Constructs an {@code ExcelException} with a detail message and a root cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause; may be {@code null}
     */
    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs an {@code ExcelException} that wraps another throwable.
     * The detail message is taken from {@code cause.toString()}.
     *
     * @param cause the underlying cause; may be {@code null}
     */
    public ExcelException(Throwable cause) {
        super(cause);
    }
}
