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
package io.github.dhsolo.common;

import java.text.SimpleDateFormat;

/**
 * General-purpose string and date utility methods used throughout the Excel framework.
 *
 * <p>This class provides lightweight helpers for null/blank checks and date formatting
 * that intentionally avoid pulling in heavyweight third-party libraries.
 *
 * @author dhsolo
 * @since 1.0
 */
public class CommonUtil {

    /**
     * Tests whether a string is {@code null} or contains only whitespace characters.
     *
     * @param object the string to test, may be {@code null}
     * @return {@code true} if {@code object} is {@code null} or blank; {@code false} otherwise
     */
    public static boolean isEmpty(String object) {
        return object == null || object.trim().length() == 0;
    }

    /**
     * Formats a date/time object into a string using the specified pattern.
     *
     * <p>If {@code pattern} is {@code null} or blank, the default pattern
     * {@code "yyyy-MM-dd HH:mm:ss"} is used.
     *
     * @param object  the date or time value to format; must be a type accepted by
     *                {@link SimpleDateFormat#format(Object)} (e.g. {@link java.util.Date})
     * @param pattern a {@link SimpleDateFormat}-compatible pattern string, or {@code null}
     *                to fall back to the default pattern
     * @return the formatted date/time string
     * @throws IllegalArgumentException if {@code object} cannot be formatted by
     *                                  {@link SimpleDateFormat}
     */
    public static String formatDate(Object object, String pattern) {
        if (isEmpty(pattern))
            pattern = "yyyy-MM-dd HH:mm:ss";
        SimpleDateFormat format = new SimpleDateFormat(pattern);
        return format.format(object);
    }
}
