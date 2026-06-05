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

/**
 * Configures image resizing for image columns in an Excel sheet.
 *
 * <p>This annotation can be used as the value of {@link ExcelInfo#imageResize()} to
 * apply a global resize policy to all image columns in the sheet. When
 * {@link #needResize()} is {@code true}, every downloaded image is scaled to the
 * specified dimensions before being embedded in the workbook.
 *
 * <pre>{@code
 * @ExcelInfo(
 *     sheetName = "Products",
 *     imageResize = @ExcelImageResize(needResize = true, resizeWidth = 120, resizeHeight = 80)
 * )
 * public class ProductExcelModel { ... }
 * }</pre>
 *
 * @author dhsolo
 * @since 1.0
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelImageResize {

    /**
     * Whether images should be resized before being embedded.
     * Defaults to {@code false} (images are embedded at their original dimensions).
     */
    boolean needResize() default false;

    /**
     * Target width in pixels to which images are scaled when {@link #needResize()} is
     * {@code true}. Defaults to {@code 500}.
     */
    int resizeWidth() default 500;

    /**
     * Target height in pixels to which images are scaled when {@link #needResize()} is
     * {@code true}. Defaults to {@code 500}.
     */
    int resizeHeight() default 500;
}
