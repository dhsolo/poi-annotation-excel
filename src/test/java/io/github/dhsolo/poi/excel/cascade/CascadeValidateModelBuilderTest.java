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
import io.github.dhsolo.poi.excel.annotation.ExcelListBox;
import io.github.dhsolo.poi.excel.core.ExcelCascadeAble;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class CascadeValidateModelBuilderTest {

    @Test
    void builder_singleColumn_noChildren() {
        CascadeValidateModel model = CascadeValidateModelBuilder.builder("alarmLevel")
                .addItem("低")
                .addItem("中")
                .addItem("高")
                .build();

        assertThat(model.getFieldName()).isEqualTo("alarmLevel");
        assertThat(model.getItems()).hasSize(3);
        assertThat(model.getItems()).extracting(CascadeValidateModel.CascadeValidateItem::getValue)
                .containsExactly("低", "中", "高");
    }

    @Test
    void builder_addItems_batch() {
        List<String> names = List.of("A", "B", "C");

        CascadeValidateModel model = CascadeValidateModelBuilder.builder("col")
                .addItems(names, s -> s)
                .build();

        assertThat(model.getItems()).hasSize(3);
    }

    @Test
    void builder_twolevel_cascade() {
        CascadeValidateModel model = CascadeValidateModelBuilder.builder("bigType")
                .addItem("模拟量",
                        CascadeValidateModelBuilder.builder("smallType")
                                .addItem("电流")
                                .addItem("电压"))
                .build();

        CascadeValidateModel.CascadeValidateItem item = model.getItems().get(0);
        assertThat(item.getValue()).isEqualTo("模拟量");
        assertThat(item.getCascadeValidateModels()).hasSize(1);
        CascadeValidateModel child = item.getCascadeValidateModels().get(0);
        assertThat(child.getFieldName()).isEqualTo("smallType");
        assertThat(child.getItems()).hasSize(2);
    }

    @Test
    void addItem_nullChildBuilder_ignored() {
        CascadeValidateModel model = CascadeValidateModelBuilder.builder("col")
                .addItem("value", (CascadeValidateModelBuilder) null)
                .build();

        CascadeValidateModel.CascadeValidateItem item = model.getItems().get(0);
        assertThat(item.getValue()).isEqualTo("value");
        assertThat(item.getCascadeValidateModels()).isNullOrEmpty();
    }

    @Test
    void hasItems_returnsTrue_afterAddItem() {
        CascadeValidateModelBuilder builder = CascadeValidateModelBuilder.builder("col");
        assertThat(builder.hasItems()).isFalse();
        builder.addItem("x");
        assertThat(builder.hasItems()).isTrue();
    }

    @Test
    void addItems_nullCollection_noop() {
        CascadeValidateModel model = CascadeValidateModelBuilder.builder("col")
                .addItems((java.util.List<String>) null, s -> s)
                .build();

        assertThat(model.getItems()).isNullOrEmpty();
    }

    @Test
    void threelevel_cascade() {
        CascadeValidateModel model = CascadeValidateModelBuilder.builder("a")
                .addItem("A1",
                        CascadeValidateModelBuilder.builder("b")
                                .addItem("B1",
                                        CascadeValidateModelBuilder.builder("c")
                                                .addItem("C1")))
                .build();

        CascadeValidateModel b = model.getItems().get(0).getCascadeValidateModels().get(0);
        CascadeValidateModel c = b.getItems().get(0).getCascadeValidateModels().get(0);
        assertThat(c.getFieldName()).isEqualTo("c");
        assertThat(c.getItems().get(0).getValue()).isEqualTo("C1");
    }

    @Test
    void TestExcelCascade(@org.junit.jupiter.api.io.TempDir java.io.File tmp){
        Cascade cascade = new Cascade();
        ExcelCreator creator = new ExcelCreator(cascade);
        creator.createExcel();
        creator.exportLocal(new java.io.File(tmp, "cascade.xlsx").getAbsolutePath());
    }

    @ExcelInfo
    class Cascade implements ExcelCascadeAble {



        @ExcelColumn(columnName = "字段1",index = 0)
        private String fieldName;

        @ExcelColumn(columnName = "字段2",index = 0)
        private String childName;

        @ExcelListBox(listTextBox = {"4","5","6"})
        @ExcelColumn(columnName = "字段3",index = 0)
        private String te;


        @Override
        public List<CascadeValidateModel> cascadeList() {
            CascadeValidateModel build = CascadeValidateModelBuilder.builder("fieldName").addItem("1").addItem("2", CascadeValidateModelBuilder.builder("childName").addItem("2-1").addItem("2-2")).build();
            return List.of(build);
        }
    }
}
