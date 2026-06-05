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
 * Marks a field as a nested child model whose columns should be flattened into the
 * parent sheet.
 *
 * <p>Apply this annotation on a field in a model class annotated with {@link ExcelInfo}
 * when the field itself holds another model object (not a list) that carries its own
 * {@link ExcelColumn}-annotated fields. The framework traverses the child object and
 * appends its columns to the parent sheet as if they were declared directly on the
 * parent class, enabling hierarchical or composed model structures without code
 * duplication.
 *
 * <pre>{@code
 * @ExcelInfo(sheetName = "Orders")
 * public class OrderExcelModel {
 *     @ExcelInfoChild
 *     private AddressInfo address;   // AddressInfo has its own @ExcelColumn fields
 * }
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelInfoChild {
}
