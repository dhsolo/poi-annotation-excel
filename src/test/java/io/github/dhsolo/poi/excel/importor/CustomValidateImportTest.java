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
package io.github.dhsolo.poi.excel.importor;

import io.github.dhsolo.poi.excel.ExcelUtil;
import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import io.github.dhsolo.poi.excel.annotation.ExcelCustomValidateMethod;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import io.github.dhsolo.poi.excel.model.ExcelRowData;
import io.github.dhsolo.poi.excel.validation.ExcelCustomValidate;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/** End-to-end: @ExcelCustomValidateMethod is invoked during import and surfaces its message. */
class CustomValidateImportTest {

    @Test
    void invalidRowTriggersCustomValidationError() throws Exception {
        byte[] xlsx = buildXlsx("年龄", "-5");   // header + one invalid row
        assertThatThrownBy(() ->
                ExcelUtil.importExcel(new ByteArrayInputStream(xlsx), 0, 1, AgeModel.class))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("年龄不能为负");
    }

    @Test
    void validRowPassesCustomValidation() throws Exception {
        byte[] xlsx = buildXlsx("年龄", "20");
        // Should not throw; the single valid row is parsed.
        assertThat(ExcelUtil.importExcel(new ByteArrayInputStream(xlsx), 0, 1, AgeModel.class)).hasSize(1);
    }

    private static byte[] buildXlsx(String header, String value) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("Sheet1");
            sheet.createRow(0).createCell(0).setCellValue(header);
            sheet.createRow(1).createCell(0).setCellValue(value);
            wb.write(out);
            return out.toByteArray();
        }
    }

    @ExcelInfo(sheetName = "age")
    public static class AgeModel {
        @ExcelColumn(columnName = "年龄", index = 1)
        private Integer age;

        @ExcelCustomValidateMethod(columnName = "age")
        public ExcelCustomValidate<AgeModel> ageRule() {
            return new ExcelCustomValidate<>() {
                @Override
                public boolean validate(ExcelRowData<AgeModel> d) {
                    Object v = d.currentValue();
                    int age = v instanceof Number ? ((Number) v).intValue()
                            : Integer.parseInt(v.toString().trim());
                    return age >= 0;
                }

                @Override
                public String errorMessage() {
                    return "年龄不能为负";
                }
            };
        }

        public Integer getAge() { return age; }
        public void setAge(Integer age) { this.age = age; }
    }
}
