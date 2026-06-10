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
package io.github.dhsolo.poi.excel.cascade;

import io.github.dhsolo.poi.excel.ExcelCreator;
import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import io.github.dhsolo.poi.excel.core.ExcelCascadeAble;
import io.github.dhsolo.poi.excel.exception.ExcelException;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.stream.Collectors;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Defined-name handling for cascade dropdowns. Excel defined names allow only
 * letters/digits/'.'/'_', are capped at 255 characters, must not look like a cell reference,
 * and are case-insensitive. Illegal characters in option values are transparently substituted
 * with '_' and mirrored into the INDIRECT formulas via SUBSTITUTE; the remaining violations
 * (overlong chains, cell-reference-like or empty values) fail fast with an
 * {@code ExcelException} instead of a bare POI error; case-variant duplicates are
 * disambiguated with a suffix instead of crashing.
 */
class CascadeNameManagerTest {

    @Test
    void illegalCharacterValueIsTransparentlySubstituted() {
        // spaces are illegal in defined names: the name is sanitised to New_York and the
        // INDIRECT formula mirrors the substitution so the lookup still resolves
        try (ExcelCreator creator = new ExcelCreator(model("New York"))) {
            creator.createExcel();
            Workbook wb = creator.getWorkBook();
            assertThat(wb.getAllNames()).extracting(Name::getNameName).containsExactly("New_York");
            assertThat(indirectFormulas(wb)).anyMatch(f -> f.contains("SUBSTITUTE") && f.contains("\" \""));
        }
    }

    @Test
    void hyphenValueIsTransparentlySubstituted() {
        try (ExcelCreator creator = new ExcelCreator(model("重庆-市区"))) {
            creator.createExcel();
            Workbook wb = creator.getWorkBook();
            assertThat(wb.getAllNames()).extracting(Name::getNameName).containsExactly("重庆_市区");
            assertThat(indirectFormulas(wb)).anyMatch(f -> f.contains("SUBSTITUTE") && f.contains("\"-\""));
        }
    }

    private static java.util.List<String> indirectFormulas(Workbook wb) {
        java.util.List<String> formulas = new java.util.ArrayList<>();
        wb.getSheetAt(0).getDataValidations().forEach(v -> {
            String f = v.getValidationConstraint().getFormula1();
            if (f != null) formulas.add(f);
        });
        return formulas;
    }

    @Test
    void overlongValueChainFailsFastWithLimit() {
        assertThatThrownBy(() -> create(model("很长的选项".repeat(60))))
                .isInstanceOf(ExcelException.class)
                .hasMessageContaining("255");
    }

    @Test
    void cellReferenceLikeValueFailsFast() {
        assertThatThrownBy(() -> create(model("A1")))
                .isInstanceOf(ExcelException.class)
                .hasMessageContaining("A1");
    }

    @Test
    void emptyValueFailsFastInsteadOfIndexOutOfBounds() {
        assertThatThrownBy(() -> create(model("")))
                .isInstanceOf(ExcelException.class)
                .hasMessageContaining("non-empty");
    }

    @Test
    void caseVariantParentValuesGetDistinctNames() {
        try (ExcelCreator creator = new ExcelCreator(new CaseModel())) {
            creator.createExcel();
            Workbook wb = creator.getWorkBook();
            Set<String> lowerNames = wb.getAllNames().stream()
                    .map(Name::getNameName)
                    .map(n -> n.toLowerCase(Locale.ROOT))
                    .collect(Collectors.toSet());
            // both parent values registered, disambiguated case-insensitively
            assertThat(wb.getAllNames()).hasSize(2);
            assertThat(lowerNames).hasSize(2);
            assertThat(lowerNames).contains("abc");
        }
    }

    @Test
    void digitFirstValueStillExportsViaUnderscorePrefix() {
        try (ExcelCreator creator = new ExcelCreator(model("2"))) {
            creator.createExcel();
            assertThat(creator.getWorkBook().getAllNames())
                    .extracting(Name::getNameName)
                    .containsExactly("_2");
        }
    }

    private static void create(Object model) {
        try (ExcelCreator creator = new ExcelCreator(model)) {
            creator.createExcel();
        }
    }

    private static CascadeModel model(String parentValue) {
        CascadeModel m = new CascadeModel();
        m.parentValue = parentValue;
        return m;
    }

    @ExcelInfo
    public static class CascadeModel implements ExcelCascadeAble {
        String parentValue;

        @ExcelColumn(columnName = "父级", index = 1)
        private String p;

        @ExcelColumn(columnName = "子级", index = 2)
        private String c;

        @Override
        public List<CascadeValidateModel> cascadeList() {
            return List.of(CascadeValidateModelBuilder.builder("p")
                    .addItem(parentValue, CascadeValidateModelBuilder.builder("c").addItem("x").addItem("y"))
                    .build());
        }
    }

    @ExcelInfo
    public static class CaseModel implements ExcelCascadeAble {
        @ExcelColumn(columnName = "父级", index = 1)
        private String p;

        @ExcelColumn(columnName = "子级", index = 2)
        private String c;

        @Override
        public List<CascadeValidateModel> cascadeList() {
            return List.of(CascadeValidateModelBuilder.builder("p")
                    .addItem("ABC", CascadeValidateModelBuilder.builder("c").addItem("x"))
                    .addItem("abc", CascadeValidateModelBuilder.builder("c").addItem("y"))
                    .build());
        }
    }
}
