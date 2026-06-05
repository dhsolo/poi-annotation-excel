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
 * Represents an Excel {@code IF} formula and provides a structured way to build
 * conditional cell expressions used in generated workbooks.
 *
 * <p>An {@code IF} formula follows the Excel syntax:
 * <pre>
 *   IF(logical_test, value_if_true, value_if_false)
 * </pre>
 * This class encapsulates the three components so that callers can compose
 * nested or chained conditions without constructing raw formula strings by hand.
 *
 * <p>Example of the formula this class is intended to produce:
 * <pre>{@code
 * // Produces: IF(C2="Y","Active","Inactive")
 * String formula = new IFFormula()
 *         .condition("C2=\"Y\"")
 *         .thenValue("\"Active\"")
 *         .elseValue("\"Inactive\"")
 *         .build();
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
public class IFFormula {
}
