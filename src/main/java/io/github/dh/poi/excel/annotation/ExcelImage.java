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
package io.github.dh.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Marks a field as an image column and configures how images are fetched and stored.
 *
 * <p>Apply this annotation together with {@link ExcelColumn} on a field whose value is
 * an image URL (or a separator-delimited list of URLs when
 * {@link ExcelInfo#imageSeparator()} is set). The framework downloads each image and
 * embeds it in the corresponding cell during export.
 *
 * <pre>{@code
 * @ExcelColumn(columnName = "Photo", index = 3)
 * @ExcelImage(imageVisitPrev = "https://cdn.example.com/",
 *             imageDownPath  = "/tmp/images/")
 * private String photoUrl;
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelImage {

    /**
     * URL prefix prepended to the field value when constructing the full image URL for
     * download. Defaults to an empty string (the field value is used as-is).
     */
    String imageVisitPrev() default "";

    /**
     * Local directory path where downloaded images are temporarily cached.
     * Defaults to an empty string (no local caching; images are streamed directly).
     */
    String imageDownPath() default "";
}
