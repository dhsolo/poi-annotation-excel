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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.ParsePosition;
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

    private static final Logger log = LoggerFactory.getLogger(CommonUtil.class);

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
            "yyyy-MM-dd HH:mm:ss.SSS",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-dd'T'HH:mm:ss",
            "yyyy/MM/dd HH:mm:ss",
            "yyyy-MM-dd HH:mm",
            "yyyy/MM/dd HH:mm",
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "yyyy-MM",
            "yyyy/MM",
            "yyyyMMdd"
    };

    /**
     * Parses a date string by trying the known patterns in order. Parsing is strict: the
     * pattern must consume the whole string (a date-only pattern never silently swallows a
     * trailing time-of-day) and field values must be in range ({@code 2024-99-99} is rejected
     * instead of leniently rolling over to a different date).
     *
     * @param source the date text; may be {@code null}
     * @return the parsed {@link Date}, or {@code null} when nothing matches (logged as WARN)
     */
    public static Date parseDate(String source) {
        return parseDate(source, null);
    }

    /**
     * Parses a date string, trying {@code preferredPattern} first (when supplied) before the
     * built-in patterns. The import pipeline passes the column's {@code @ExcelDateFormat}
     * pattern here so a value that was formatted with a non-standard pattern (e.g.
     * {@code dd-MM-yyyy}) round-trips back to a {@link Date} instead of silently becoming
     * {@code null}.
     *
     * @param source          the date text; may be {@code null}
     * @param preferredPattern a {@link SimpleDateFormat} pattern to try first; may be
     *                         {@code null}/blank to use only the built-in patterns
     * @return the parsed {@link Date}, or {@code null} when nothing matches (logged as WARN)
     */
    public static Date parseDate(String source, String preferredPattern) {
        if (source == null) return null;
        String trimmed = source.trim();
        if (trimmed.isEmpty()) return null;
        if (preferredPattern != null && !preferredPattern.trim().isEmpty()) {
            Date parsed = tryParse(trimmed, preferredPattern);
            if (parsed != null) return parsed;
        }
        for (String pattern : DATE_PATTERNS) {
            Date parsed = tryParse(trimmed, pattern);
            if (parsed != null) return parsed;
        }
        log.warn("Unable to parse date string '{}' with any of the known patterns", trimmed);
        return null;
    }

    /**
     * Strictly parses {@code trimmed} against a single {@code pattern}: the pattern must consume
     * the whole string and field values must be in range. Returns {@code null} on any mismatch
     * (including an invalid pattern string), so callers can simply try the next candidate.
     */
    private static Date tryParse(String trimmed, String pattern) {
        try {
            SimpleDateFormat format = new SimpleDateFormat(pattern);
            format.setLenient(false);
            ParsePosition pos = new ParsePosition(0);
            Date parsed = format.parse(trimmed, pos);
            if (parsed != null && pos.getIndex() == trimmed.length()) {
                return parsed;
            }
        } catch (IllegalArgumentException invalidPattern) {
            // malformed pattern string; treat as a non-match and fall through to the next one
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
        return convert(value, type, null);
    }

    /**
     * Variant of {@link #convert(Object, Class)} that lets the caller supply the date pattern the
     * value was originally formatted with, so date/time targets re-parse with that pattern first
     * (see {@link #parseDate(String, String)}). All integral targets ({@code Integer}, {@code Long},
     * {@code Short}, {@code Byte}, {@code BigInteger}) share the same truncate-toward-zero semantics
     * via {@link #integerPart}, so a fractional cell maps consistently regardless of the field's
     * integral type instead of one type truncating and another throwing.
     *
     * @param value       the raw value to convert; may be {@code null}
     * @param type        the target type; may be {@code null} (returns {@code null})
     * @param datePattern the preferred date pattern for date/time targets; may be {@code null}
     * @return the converted value, or {@code null} when not convertible
     */
    public static Object convert(Object value, Class<?> type, String datePattern) {
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
            result = integerPart(numVal).intValueExact();
            break;
        case "java.lang.Double":
        case "double":
            result = Double.valueOf(numVal);
            break;
        case "java.lang.Long":
        case "long":
            result = integerPart(numVal).longValueExact();
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
            result = integerPart(numVal).shortValueExact();
            break;
        case "java.util.Date":
            result = parseDate(value.toString(), datePattern);
            break;
        case "java.sql.Timestamp": {
            Date d = parseDate(value.toString(), datePattern);
            result = d != null ? new java.sql.Timestamp(d.getTime()) : null;
            break;
        }
        case "java.sql.Date": {
            Date d = parseDate(value.toString(), datePattern);
            result = d != null ? new java.sql.Date(d.getTime()) : null;
            break;
        }
        case "java.math.BigDecimal":
            result = new java.math.BigDecimal(numVal);
            break;
        case "java.math.BigInteger":
            result = integerPart(numVal);
            break;
        case "java.lang.Character":
        case "char": {
            String charStr = value.toString();
            result = charStr.isEmpty() ? null : Character.valueOf(charStr.charAt(0));
            break;
        }
        case "java.lang.Byte":
        case "byte":
            result = integerPart(numVal).byteValueExact();
            break;
        case "java.time.LocalDate": {
            Date d = parseDate(value.toString(), datePattern);
            result = d != null ? d.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate() : null;
            break;
        }
        case "java.time.LocalDateTime": {
            Date d = parseDate(value.toString(), datePattern);
            result = d != null ? d.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDateTime() : null;
            break;
        }
        }
        return result;
    }

    /**
     * The integer part of a numeric string, truncated toward zero (legacy import semantics
     * for decimal cell text). Goes through {@link java.math.BigDecimal} so scientific
     * notation resolves exactly — a naive cut at the first {@code '.'} would turn
     * {@code "1.2345678E7"} (the {@code Double.toString} form of 12345678) into {@code 1}.
     *
     * @throws NumberFormatException when the string is not numeric
     */
    private static java.math.BigInteger integerPart(String numVal) {
        return new java.math.BigDecimal(numVal).toBigInteger();
    }
}
