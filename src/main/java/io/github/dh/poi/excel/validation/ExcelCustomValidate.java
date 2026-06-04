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
package io.github.dh.poi.excel.validation;

import io.github.dh.poi.excel.model.ExcelRowData;

/**
 * Strategy interface for applying custom row-level validation logic during Excel import.
 *
 * <p>Implement this interface to enforce business rules that go beyond simple type conversion
 * or non-null checks. The framework invokes {@link #validate(ExcelRowData)} for each data row
 * after the raw cell values have been read and coerced to the target type {@code T}. If
 * validation fails the framework rejects the row and records {@link #errorMessage()} in
 * the import result, without aborting the entire import.
 *
 * <p>Example — reject a row whose quantity field is negative:
 * <pre>{@code
 * public class PositiveQtyValidator implements ExcelCustomValidate<OrderRow> {
 *     public boolean validate(ExcelRowData<OrderRow> data) {
 *         return data.getRowData().getQuantity() > 0;
 *     }
 *     public String errorMessage() {
 *         return "Quantity must be greater than zero";
 *     }
 * }
 * }</pre>
 *
 * @param <T> the type of the model object that each row of the Excel file is mapped to
 * @author dh
 * @since 1.0
 */
public interface ExcelCustomValidate<T> {

    /**
     * Determines whether the given row passes this validation rule.
     *
     * <p>This method is called once per data row. Returning {@code false} causes the framework
     * to mark the row as invalid and skip it from the successful import results.
     *
     * @param excelRowData the context object for the current row, providing access to the
     *                     typed model object, raw cell values, row index, and POI primitives
     * @return {@code true} if the row is valid and should be included in the import results;
     *         {@code false} if the row should be rejected
     */
    boolean validate(ExcelRowData<T> excelRowData);

    /**
     * Returns a human-readable message that describes why a row failed this validation rule.
     *
     * <p>The message is attached to rejected rows in the import result so that end-users
     * can understand which business rule was violated and correct the data in the source file.
     *
     * @return a non-null, non-empty validation error message
     */
    String errorMessage();
}
