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
package io.github.dhsolo.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {
	String columnName() default "";
	int index();
	boolean nullable() default true;
	int columnWidth() default 20;
	boolean needMergeCell() default false;
	int mergeCellIndex() default 1;
	String noneCellDefaultValue() default "";

	/**
	 * Inline enum translation in "key:displayValue" format (e.g. {@code {"0:Pending","1:Approved"}}).
	 * <p>Avoids writing a corresponding {@code @ExcelTranslateMethod} method; suitable for simple int/String → text mappings.
	 * <p>If {@code @ExcelTranslateMethod} also exists on the field, the method translation takes precedence.
	 */
	String[] translate() default {};

	/**
	 * Ancestor header labels for a multi-level (3+ row) merged header, ordered top-most first;
	 * this column's {@link #columnName()} is the bottom (leaf) header. Adjacent columns that share
	 * the same ancestor prefix are merged horizontally at each level, and a column with fewer
	 * levels than the deepest one has its leaf merged vertically down to the bottom row.
	 *
	 * <p>Example — a two-level group above the leaf produces a three-row header:
	 * <pre>{@code
	 * @ExcelColumn(columnName = "Q1", index = 2, groups = {"2024年", "上半年"})
	 * @ExcelColumn(columnName = "Q2", index = 3, groups = {"2024年", "上半年"})
	 * @ExcelColumn(columnName = "Q3", index = 4, groups = {"2024年", "下半年"})
	 * }</pre>
	 * renders {@code 2024年} spanning Q1–Q3 on the top row, {@code 上半年}/{@code 下半年} on the
	 * middle row, and {@code Q1/Q2/Q3} on the bottom row. Arbitrary depth is supported.
	 *
	 * <p>Mutually exclusive with {@link ExcelColumnParent} (the legacy two-row mechanism) on the
	 * same model; mixing the two throws at processing time. Export-only, like {@code @ExcelColumnParent}.
	 */
	String[] groups() default {};

	/**
	 * Dot-separated path: reads a value from the row data object by path, replacing nested-object {@code @ExcelTranslateMethod}.
	 * <p>For example, {@code "pointInfo.measureType"} is equivalent to {@code data.getPointInfo().getMeasureType()}.
	 * <p>Can be combined with {@link #translate()}; if {@code @ExcelTranslateMethod} also exists, method translation takes precedence.
	 */
	String sourcePath() default "";

	/**
	 * Same-level field redirect: reads another field of the row data object to fill this column,
	 * suitable when the display column name differs from the value field name.
	 * <p>For example, {@code sourceField = "correctResult"} means this column displays the value of {@code correctResult}.
	 * <p>Can be combined with {@link #translate()}; if {@code @ExcelTranslateMethod} also exists, method translation takes precedence.
	 */
	String sourceField() default "";
}
