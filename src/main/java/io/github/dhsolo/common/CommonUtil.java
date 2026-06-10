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

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

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

    /** Date patterns tried, most specific first, when parsing date strings. */
    private static final String[] DATE_PATTERNS = {
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-dd'T'HH:mm:ss",
            "yyyy/MM/dd HH:mm:ss",
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "yyyy-MM",
            "yyyy/MM",
            "yyyyMMdd"
    };

    /**
     * Parses a date string by trying the known patterns in order.
     *
     * @param source the date text; may be {@code null}
     * @return the parsed {@link Date}, or {@code null} when nothing matches
     */
    public static Date parseDate(String source) {
        if (source == null) return null;
        String trimmed = source.trim();
        for (String pattern : DATE_PATTERNS) {
            try {
                return new SimpleDateFormat(pattern).parse(trimmed);
            } catch (ParseException ignored) {}
        }
        return null;
    }

    /**
     * Converts a raw (typically String) value to the requested target type. Supports the
     * boxed/primitive numerics, {@code String}, {@code Boolean}, {@code Character},
     * {@code BigDecimal}/{@code BigInteger}, and the common date types. Returns {@code null}
     * when the value is blank or the target type is unsupported.
     *
     * <p>Shared by the import pipeline ({@code ExcelImportor.caseObject}) and
     * {@code Reflect.mapToBean}, so row-to-bean mapping applies the same conversions as
     * column imports.
     *
     * @param value the raw value to convert; may be {@code null}
     * @param type  the target type; may be {@code null} (returns {@code null})
     * @return the converted value, or {@code null} when not convertible
     */
    public static Object convert(Object value, Class<?> type) {
        if (type == null || value == null || value.toString().trim().length() == 0)
            return null;
        Object result = null;
        String className = type.getCanonicalName();
        String numVal = value.toString().replaceAll(",", "").trim();
        switch (className) {
        case "java.lang.String":
            result = value + "";
            break;
        case "java.lang.Integer":
        case "int":
            if (numVal.indexOf(".") != -1) {
                numVal = numVal.substring(0, numVal.indexOf("."));
            }
            result = Integer.valueOf(numVal);
            break;
        case "java.lang.Double":
        case "double":
            result = Double.valueOf(numVal);
            break;
        case "java.lang.Long":
        case "long":
            result = Long.valueOf(numVal);
            break;
        case "java.lang.Boolean":
        case "boolean":
            result = Boolean.valueOf(value.toString());
            break;
        case "java.lang.Float":
        case "float":
            result = Float.valueOf(numVal);
            break;
        case "java.lang.Short":
        case "short":
            result = Short.valueOf(numVal);
            break;
        case "java.util.Date":
            result = parseDate(value.toString());
            break;
        case "java.sql.Timestamp": {
            Date d = parseDate(value.toString());
            result = d != null ? new java.sql.Timestamp(d.getTime()) : null;
            break;
        }
        case "java.sql.Date": {
            Date d = parseDate(value.toString());
            result = d != null ? new java.sql.Date(d.getTime()) : null;
            break;
        }
        case "java.math.BigDecimal":
            result = new java.math.BigDecimal(numVal);
            break;
        case "java.math.BigInteger":
            result = new java.math.BigInteger(numVal.indexOf(".") != -1
                    ? numVal.substring(0, numVal.indexOf("."))
                    : numVal);
            break;
        case "java.lang.Character":
        case "char": {
            String charStr = value.toString();
            result = charStr.isEmpty() ? null : Character.valueOf(charStr.charAt(0));
            break;
        }
        case "java.lang.Byte":
        case "byte":
            result = Byte.valueOf(numVal);
            break;
        case "java.time.LocalDate": {
            Date d = parseDate(value.toString());
            result = d != null ? d.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate() : null;
            break;
        }
        case "java.time.LocalDateTime": {
            Date d = parseDate(value.toString());
            result = d != null ? d.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDateTime() : null;
            break;
        }
        }
        return result;
    }
}
