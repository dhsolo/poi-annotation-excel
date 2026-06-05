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
package io.github.dhsolo.poi.excel.formula;

/**
 * Fluent builder for constructing Excel formula strings to be written into cells
 * via the {@code @ExcelFormula} annotation or the framework's cell-value pipeline.
 *
 * <p>Implementations of this builder are expected to accumulate formula fragments
 * and produce a final formula string that conforms to the OpenXML / Apache POI
 * formula syntax (e.g. {@code "SUM(A1:A10)"}, {@code "IF(B2>0,\"yes\",\"no\")"}).
 *
 * <p>Typical usage pattern:
 * <pre>{@code
 * String formula = new FormulaBuilder()
 *         .sum("A1", "A10")
 *         .build();
 * }</pre>
 *
 * @author dhsolo
 * @since 1.0
 */
public class FormulaBuilder {
}
