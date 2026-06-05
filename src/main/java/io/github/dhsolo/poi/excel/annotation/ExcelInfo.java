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

import io.github.dhsolo.poi.excel.ExcelCreator;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Type-level annotation that controls the overall Excel export / import configuration
 * for a model class.
 *
 * <p>Place this annotation on the model class that represents a single Excel sheet.
 * All other annotations ({@link ExcelColumn}, {@link ExcelData}, {@link ExcelTitle}, etc.)
 * are discovered from the same class at runtime.
 *
 * <pre>{@code
 * @ExcelInfo(sheetName = "Users", excelType = ExcelCreator.XLSX, needOrder = true)
 * public class UserExcelModel {
 *     @ExcelTitle
 *     private String title = "User Report";
 *
 *     @ExcelData
 *     private List<UserRow> rows;
 *     // ...
 * }
 * }</pre>
 *
 * @author dhsolo
 * @since 1.0
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelInfo {

	/**
	 * Zero-based sheet index (or indices) to read from when importing.
	 * Defaults to {@code 0} (the first sheet).
	 */
	int[] sheetNum() default 0;

	/**
	 * Output file type: {@code "xlsx"} (default) or {@code "xls"}.
	 * Use the constants {@link ExcelCreator#XLSX} / {@code ExcelCreator.XLS}.
	 */
	String excelType() default ExcelCreator.XLSX;

	/**
	 * Whether to prepend an auto-incremented sequence-number column as the first column.
	 * Defaults to {@code false}.
	 */
	boolean needOrder() default false;

	/**
	 * Number of columns the sequence-number column should span horizontally (merge width).
	 * When greater than {@code 1}, the "序号" column is widened to cover this many columns
	 * and they are merged into a single cell on each row.
	 * Only meaningful when {@link #needOrder()} is {@code true}. Defaults to {@code 1}.
	 */
	int orderColumnSpan() default 1;

	/**
	 * Row height of the title row in POI units (1/20 of a point).
	 * Defaults to {@code 2000}.
	 */
	int titleHeight() default 2000;

	/**
	 * Row height of the header row in POI units (1/20 of a point).
	 * Defaults to {@code 2000}.
	 */
	int headerHeight() default 2000;

	/**
	 * Zero-based row index at which data rows begin when importing.
	 * Defaults to {@code 0}.
	 */
	int startRow() default 0;

	/**
	 * Zero-based column indices to skip during import.
	 * Defaults to an empty array (no columns skipped).
	 */
	int[] exceptColumnNum() default {};

	/**
	 * Name of the sheet tab in the generated workbook.
	 * Defaults to an empty string, in which case the framework uses a default sheet name.
	 */
	String sheetName() default "";

	/**
	 * Timeout in milliseconds for downloading remote images during export.
	 * Defaults to {@code 500}.
	 */
	int imageReadTimeOut() default 500;

	/**
	 * Separator used when a cell contains multiple image URLs.
	 * Defaults to an empty string (single image per cell).
	 */
	String imageSeparator() default "";

	/**
	 * Anchor type for embedded pictures.
	 * Use {@link ExcelCreator#MOVE_AND_RESIZE} ({@code 0}, default) or
	 * {@code ExcelCreator.DONT_MOVE_AND_RESIZE} ({@code 3}).
	 */
	int pictureInnerType() default ExcelCreator.MOVE_AND_RESIZE;

	/**
	 * Whether to use the big-data (streaming) workbook strategy ({@code SXSSFWorkbook})
	 * for export, which reduces memory usage for large datasets.
	 * Defaults to {@code true}.
	 */
	boolean isBigData() default true;

	/**
	 * Image resize settings applied globally to all image columns in this sheet.
	 * Defaults to no resizing ({@code needResize = false}).
	 *
	 * @see ExcelImageResize
	 */
	ExcelImageResize imageResize() default @ExcelImageResize;

	/**
	 * Default string to write into a cell when the corresponding field value is {@code null}
	 * or empty during export. Defaults to an empty string.
	 */
	String noneCellDefaultValue() default "";
}
