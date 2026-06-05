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
package io.github.dhsolo.poi.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A child sheet built independently (its own workbook) keeps its custom cell style when stitched
 * into a parent workbook as an additional sheet — the style is cloned into the shared workbook
 * rather than discarded or rejected with "Style does not belong to the supplied Workbook".
 */
class MultiSheetCustomStyleTest {

    @Test
    void childSheetKeepsCustomDataStyleInSharedWorkbook() throws Exception {
        // Child built against its OWN workbook, with a custom yellow data-cell fill.
        ExcelCreator child = ExcelCreatorBuilder.create("子表")
                .columns("名称:name")
                .data(List.of(Map.of("name", "张三")))
                .build();
        CellStyle yellow = child.createCellStyle();
        yellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        yellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        child.setCellStyle(yellow);

        LinkedList<ExcelCreator> children = new LinkedList<>();
        children.add(child);

        // Parent (its own workbook) adopts the child as a second sheet and exports.
        byte[] bytes = ExcelCreatorBuilder.create("主表")
                .columns("编号:id")
                .data(List.of(Map.of("id", "A001")))
                .child(children)
                .toBytes();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            Sheet childSheet = wb.getSheet("子表");
            assertThat(childSheet).isNotNull();
            CellStyle dataStyle = childSheet.getRow(1).getCell(0).getCellStyle();
            // The custom fill survived being cloned into the parent's workbook.
            assertThat(dataStyle.getFillForegroundColor()).isEqualTo(IndexedColors.YELLOW.getIndex());
            assertThat(dataStyle.getFillPattern()).isEqualTo(FillPatternType.SOLID_FOREGROUND);
        }
    }
}
