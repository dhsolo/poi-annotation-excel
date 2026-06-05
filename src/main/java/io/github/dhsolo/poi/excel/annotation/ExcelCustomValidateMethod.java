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
 * Marks a method that provides custom import validation logic for a specific column.
 *
 * <p>Apply this annotation on a method in the model class to plug in bespoke validation
 * beyond what the built-in list-box or type checks offer. The framework invokes the
 * method during import processing for the column identified by {@link #columnName()}.
 * The method should throw an appropriate exception (or return a validation result) to
 * signal invalid cell values.
 *
 * <pre>{@code
 * @ExcelCustomValidateMethod(columnName = "age")
 * public void validateAge(Object value) {
 *     int age = Integer.parseInt(value.toString());
 *     if (age < 0 || age > 150) throw new IllegalArgumentException("Invalid age: " + age);
 * }
 * }</pre>
 *
 * @author dhsolo
 * @since 1.0
 */
@Target({ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCustomValidateMethod {

    /**
     * Name of the column field (i.e. the Java field name of the corresponding
     * {@link ExcelColumn}-annotated field) this validation method applies to.
     * This attribute is required.
     */
    String columnName();
}
