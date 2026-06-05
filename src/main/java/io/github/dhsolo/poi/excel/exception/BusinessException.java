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
 * Legacy general-purpose business exception that was previously used as the root exception
 * type for this framework.
 *
 * <p>This class has been superseded by {@link ExcelException}, which provides a more
 * descriptive name and a richer constructor set. {@code BusinessException} is retained
 * solely for binary and source compatibility with existing callers; new code should
 * throw and catch {@link ExcelException} or one of its specific subclasses instead.
 *
 * @author dh
 * @since 1.0
 * @deprecated Use {@link ExcelException} or a specific subclass instead.
 *             This class will be removed in a future major release.
 */
@Deprecated
public class BusinessException extends ExcelException {

    /**
     * Constructs a {@code BusinessException} with no detail message.
     */
    public BusinessException() {
        super();
    }

    /**
     * Constructs a {@code BusinessException} with the given detail message.
     *
     * @param message the detail message
     */
    public BusinessException(String message) {
        super(message);
    }

    /**
     * Constructs a {@code BusinessException} whose detail message is built from a format
     * string and arguments via {@link String#format(String, Object[])}.
     *
     * @param format a {@link java.util.Formatter} format string
     * @param args   arguments referenced by the format specifiers
     */
    public BusinessException(String format, Object... args) {
        super(format, args);
    }

    /**
     * Constructs a {@code BusinessException} carrying a numeric error code and a
     * human-readable message. The detail message is serialised as a compact JSON string.
     *
     * @param errorCode numeric error code (e.g. an HTTP status or application error code)
     * @param message   human-readable error description
     */
    public BusinessException(int errorCode, String message) {
        super(errorCode, message);
    }

    /**
     * Constructs a {@code BusinessException} with a detail message and root cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause; may be {@code null}
     */
    public BusinessException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs a {@code BusinessException} wrapping another throwable.
     *
     * @param cause the underlying cause; may be {@code null}
     */
    public BusinessException(Throwable cause) {
        super(cause);
    }
}
