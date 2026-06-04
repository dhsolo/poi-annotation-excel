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

import io.github.dh.poi.excel.ExcelModel;
import io.github.dh.poi.excel.model.ExcelRowData;
import io.github.dh.poi.excel.validation.ExcelCustomValidate;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/** Verifies @ExcelCustomValidateMethod is wired onto the column's ExcelModel during processing. */
class CustomValidateMethodTest {

    @Test
    void customValidateMethodIsAppliedToColumnModel() {
        DefaultAnnotationProcessor processor = new DefaultAnnotationProcessor(new Model());
        List<ExcelModel> models = processor.getExcelAnnotationProperty().getExcelModels();

        ExcelModel ageModel = models.stream()
                .filter(m -> "age".equals(m.getFieldName()))
                .findFirst().orElseThrow();

        ExcelCustomValidate<?> validator = ageModel.getExcelCustomValidate();
        assertThat(validator).isNotNull();
        assertThat(validator.errorMessage()).isEqualTo("年龄不能为负");
    }

    @ExcelInfo(sheetName = "v")
    public static class Model {
        @ExcelData
        private List<Object> data = List.of();

        @ExcelColumn(columnName = "年龄", index = 1)
        private Integer age;

        @ExcelCustomValidateMethod(columnName = "age")
        public ExcelCustomValidate<Object> ageRule() {
            return new ExcelCustomValidate<>() {
                @Override public boolean validate(ExcelRowData<Object> d) { return true; }
                @Override public String errorMessage() { return "年龄不能为负"; }
            };
        }

        public List<Object> getData() { return data; }
        public Integer getAge() { return age; }
        public void setAge(Integer age) { this.age = age; }
    }
}
