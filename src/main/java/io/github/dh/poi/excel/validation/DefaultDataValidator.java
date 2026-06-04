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

import io.github.dh.poi.excel.ExcelModel;
import io.github.dh.poi.excel.cascade.CascadeValidateItemWrapper;
import io.github.dh.poi.excel.cascade.CascadeValidateModelWrapper;
import io.github.dh.poi.excel.exception.ExcelColumnNotFoundException;
import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import io.github.dh.common.Reflect;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Default implementation of {@link DataValidator}.
 *
 * <p>Handles dropdown lists, cascade validation, formula validation, etc.
 *
 * @author dh
 * @since 1.0
 */
public class DefaultDataValidator implements DataValidator {

    private static final String XLSX = "xlsx";

    private final Workbook book;
    private final Sheet sheet;
    private final String currentExcelType;
    private final Sheet hiddenSheetListBox;
    private final boolean isBigData;

    private final Set<String> existNamaManager;
    private final AtomicInteger atomicInteger;
    private final Map<String, ExcelModel> columnNameModelMappingInfo;
    private int currentListNum;
    private int rowNum;

    private DataValidationConstraint listBoxValidate;

    /**
     * columnIndex -> columnName -> formal : addressList
     */
    private final Map<Integer, Map<String, Map<String, CellRangeAddressList>>> existINDIRECT = new HashMap<>();

    public DefaultDataValidator(Workbook book, Sheet sheet, String currentExcelType, Sheet hiddenSheetListBox,
                       boolean isBigData, Set<String> existNamaManager, AtomicInteger atomicInteger,
                       Map<String, ExcelModel> columnNameModelMappingInfo) {
        this.book = book;
        this.sheet = sheet;
        this.currentExcelType = currentExcelType;
        this.hiddenSheetListBox = hiddenSheetListBox;
        this.isBigData = isBigData;
        this.existNamaManager = existNamaManager;
        this.atomicInteger = atomicInteger;
        this.columnNameModelMappingInfo = columnNameModelMappingInfo;
    }

    @Override
    public void setCurrentListNum(int currentListNum) {
        this.currentListNum = currentListNum;
    }

    @Override
    public int getCurrentListNum() {
        return currentListNum;
    }

    @Override
    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    /**
     * Performs data validation for all columns (dropdowns, formulas, etc.)
     */
    @Override
    public void checkListBox(Map<Integer, ExcelModel> columnMappingInfo, boolean needOrderNum, int rowNum) {
        this.rowNum = rowNum;
        int initNum = needOrderNum ? 1 : 0;

        columnMappingInfo.forEach((index, model) -> {
            if (model.isListBox()) {
                CascadeValidateModelWrapper cascadeValidateModel = model.getCascadeValidateModel();
                paddingSheetListBox(cascadeValidateModel, false);
            } else if (Reflect.hasText(model.getStrFormula())) {
                setFormula(model.getStrFormula(), model.getRealIndex());
            }
        });

        if (!existINDIRECT.isEmpty()) {
            processINDIRECTValidations();
        }
    }

    private void processINDIRECTValidations() {
        Collection<Map<String, Map<String, CellRangeAddressList>>> values = existINDIRECT.values();
        for (Map<String, Map<String, CellRangeAddressList>> map : values) {
            CellRangeAddressList cellRangeAddressList = null;
            Stack<String> stack = new Stack<>();
            for (Map.Entry<String, Map<String, CellRangeAddressList>> entry : map.entrySet()) {
                Map<String, CellRangeAddressList> value = entry.getValue();
                for (Map.Entry<String, CellRangeAddressList> en : value.entrySet()) {
                    String formal = en.getKey();
                    stack.push(formal);
                    cellRangeAddressList = en.getValue();
                }
            }
            String sb = "INDIRECT(";
            boolean isFirst = true;
            while (!stack.isEmpty()) {
                String pop = stack.pop();
                if (isFirst) {
                    sb += pop;
                    isFirst = false;
                } else {
                    sb = sb.replace("@@", pop);
                }
            }
            sb += ")";
            DataValidationHelper help = createDataValidationHelper();
            listBoxValidate = help.createFormulaListConstraint(sb);
            DataValidation validation = help.createValidation(listBoxValidate, cellRangeAddressList);
            validation.setEmptyCellAllowed(false);
            if (validation instanceof XSSFDataValidation) {
                validation.setSuppressDropDownArrow(true);
                validation.setShowErrorBox(true);
            } else {
                validation.setSuppressDropDownArrow(false);
            }
            validation.createPromptBox("Dropdown Tip", "Please select a value from the dropdown list.");
            sheet.addValidationData(validation);
        }
    }

    /**
     * Sets formula into a column
     */
    private void setFormula(String strFormula, int index) {
        String finalStrFormula = strFormula;
        Map<String, String> columnCharMap = new HashMap<>();
        while (strFormula.contains("@Column")) {
            int ic = strFormula.indexOf("@Column");
            String prefix = strFormula.substring(0, ic);
            strFormula = strFormula.substring(ic);
            int startI = strFormula.indexOf("(");
            int endI = strFormula.indexOf(")");
            String columnName = strFormula.substring(startI + 1, endI).trim();
            String substring = strFormula.substring(endI + 1);
            ExcelModel excelModel = columnNameModelMappingInfo.get(columnName);
            if (excelModel == null) {
                throw new ExcelColumnNotFoundException("Formula '" + finalStrFormula + "' references an undefined column name: " + columnName);
            }
            String charByNum = transCharByNum(excelModel.getRealIndex());
            String columnChar = "@Trans(" + charByNum + ")";
            columnCharMap.put(columnChar, charByNum);
            strFormula = prefix + columnChar + substring;
        }
        String finalStrFormula1 = strFormula;
        for (int i = 0; i < 1000; i++) {
            int i1 = rowNum + i;
            // Prefer reusing an existing row via getRow to avoid overwriting data rows with createRow
            Row row1 = sheet.getRow(i1) != null ? sheet.getRow(i1) : sheet.createRow(i1);
            Cell cell1 = row1.createCell(index);
            for (Map.Entry<String, String> entry : columnCharMap.entrySet()) {
                strFormula = strFormula.replace(entry.getKey(), entry.getValue() + (i1 + 1));
            }
            cell1.setCellFormula(strFormula);
            strFormula = finalStrFormula1;
        }
    }

    public void createCellListBoxFormal(int firstRow, int lastRow, int firstCol, int lastCol,
                                 String normalFormula, boolean isNeedAddTranslationException) {
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidationHelper help = createDataValidationHelper();
        listBoxValidate = help.createFormulaListConstraint(normalFormula);
        DataValidation validation = help.createValidation(listBoxValidate, regions);
        if (isNeedAddTranslationException) {
            validation.createErrorBox("Invalid Value", "Please select a valid value from the list.");
            validation.setShowErrorBox(true);
        }
        validation.setSuppressDropDownArrow(true);
        sheet.addValidationData(validation);
    }

    public String createCellListBoxFormal(int firstRow, int lastRow, int firstCol, int lastCol,
                                   CascadeValidateModelWrapper cascadeValidateModelWrapper, String pddStr) {
        return generateFormal(firstRow + 1, cascadeValidateModelWrapper, pddStr);
    }

    private String generateFormal(int firstRow, CascadeValidateModelWrapper cascadeValidateModelWrapper, String pddStr) {
        Stack<CascadeValidateItemWrapper> stack = new Stack<>();
        CascadeValidateItemWrapper parentValue = cascadeValidateModelWrapper.getParentValue();
        while (parentValue != null) {
            stack.add(parentValue);
            parentValue = parentValue.getOwnModel().getParentValue();
        }
        StringBuilder sb = new StringBuilder();
        if (stack.size() == 1) {
            CascadeValidateItemWrapper pop = stack.pop();
            CascadeValidateModelWrapper ownModel = pop.getOwnModel();
            ExcelModel excelModel = ownModel.getExcelModel();
            if (pop.isAppendPrefix()) {
                sb.append("\"_\"").append("&").append(transCharByNum(excelModel.getRealIndex())).append(firstRow);
            } else {
                sb.append(transCharByNum(excelModel.getRealIndex())).append(firstRow);
            }
            if (Reflect.hasText(pddStr)) {
                sb.append("&\"").append(pddStr).append("\"");
            }
        } else {
            while (!stack.isEmpty()) {
                CascadeValidateItemWrapper pop = stack.pop();
                CascadeValidateModelWrapper ownModel = pop.getOwnModel();
                ExcelModel excelModel = ownModel.getExcelModel();
                String charByNum = transCharByNum(excelModel.getRealIndex());
                if (pop.isAppendPrefix()) {
                    // Use & to concatenate the "_" prefix — comma is not a valid Excel formula concatenation operator
                    sb.append("\"_\"").append("&").append("IF(").append(charByNum).append(firstRow)
                            .append("=\"\",\"null\",").append(charByNum).append(firstRow).append(")").append("&\"\"&");
                } else {
                    sb.append("IF(").append(charByNum).append(firstRow)
                            .append("=\"\",\"null\",").append(charByNum).append(firstRow).append(")").append("&\"\"&");
                }
            }
            sb = new StringBuilder(sb.substring(0, sb.length() - 4));
            if (Reflect.hasText(pddStr)) {
                sb.append("&\"").append(pddStr).append("\"");
            }
        }
        return sb.toString();
    }

    private void paddingSheetListBox(CascadeValidateModelWrapper cascadeValidateModel, boolean needCreateName) {
        List<CascadeValidateItemWrapper> items = cascadeValidateModel.getItems();
        int i = 0;
        int start = currentListNum;
        int end = currentListNum + items.size();
        List<CascadeValidateItemWrapper> children = new ArrayList<>();
        for (; currentListNum < end; currentListNum++) {
            Row row = hiddenSheetListBox.createRow(currentListNum - 1);
            Cell cell = row.createCell(0);
            CascadeValidateItemWrapper cascadeValidateItem = items.get(i++);
            cell.setCellValue(cascadeValidateItem.getValue());
            if (cascadeValidateItem.getCascadeValidateModels() != null && !cascadeValidateItem.getCascadeValidateModels().isEmpty()) {
                children.add(cascadeValidateItem);
            }
        }
        currentListNum++;
        ExcelModel model = cascadeValidateModel.getExcelModel();
        int realIndex = model.getRealIndex();
        String format = String.format("listConstantData!$A$%s:$A$%s", start, currentListNum - 2);
        if (cascadeValidateModel.getParentValue() == null) {
            model.setStrFormula(format);
            createCellListBoxFormal(rowNum, 1000, realIndex, realIndex, model.getStrFormula(),
                    cascadeValidateModel.isNeedAddTranslationException());
        }

        if (needCreateName) {
            String parentValuePath = findParentValuePath(cascadeValidateModel);
            if (!Character.isLetter(parentValuePath.charAt(0))) {
                parentValuePath = "_" + parentValuePath;
                setParentIsAppendPrefix(cascadeValidateModel);
            }
            String pddStr = null;
            if (existNamaManager.contains(parentValuePath)) {
                pddStr = getNextChar();
                parentValuePath += pddStr;
            }
            existNamaManager.add(parentValuePath);
            Name name = book.createName();
            name.setNameName(parentValuePath);
            name.setRefersToFormula(format);
            if (!existINDIRECT.containsKey(realIndex)) {
                CellRangeAddressList regions = new CellRangeAddressList(rowNum, 1000, realIndex, realIndex);
                String cellListBoxFormal = createCellListBoxFormal(rowNum, 1000, realIndex, realIndex, cascadeValidateModel, pddStr);
                Map<String, Map<String, CellRangeAddressList>> stringMapMap = existINDIRECT.computeIfAbsent(realIndex, k -> new LinkedHashMap<>());
                Map<String, CellRangeAddressList> stringCellRangeAddressListMap = stringMapMap.computeIfAbsent(
                        cascadeValidateModel.getParentValue().getOwnModel().getExcelModel().getFieldName(), k -> new LinkedHashMap<>());
                stringCellRangeAddressListMap.put(cellListBoxFormal, regions);
            } else {
                Map<String, Map<String, CellRangeAddressList>> stringMapMap = existINDIRECT.get(realIndex);
                CellRangeAddressList regions = new CellRangeAddressList(rowNum, 1000, realIndex, realIndex);
                String cellListBoxFormal = generateFormal(rowNum + 1, cascadeValidateModel, pddStr);
                cellListBoxFormal = String.format("IF(%s=\"%s\",%s,@@)", cellListBoxFormal, parentValuePath, cellListBoxFormal);
                Map<String, CellRangeAddressList> stringCellRangeAddressListMap = stringMapMap.computeIfAbsent(
                        cascadeValidateModel.getParentValue().getOwnModel().getExcelModel().getFieldName(), k -> new LinkedHashMap<>());
                stringCellRangeAddressListMap.put(cellListBoxFormal, regions);
            }
        }
        if (!children.isEmpty()) {
            for (CascadeValidateItemWrapper item : children) {
                for (CascadeValidateModelWrapper wrapper : item.getCascadeValidateModels()) {
                    paddingSheetListBox(wrapper, true);
                }
            }
        }
    }

    private String getNextChar() {
        return transCharByNum(atomicInteger.getAndIncrement());
    }

    static String transCharByNum(Integer num) {
        int divResult = num / 26;
        int remainderResult = num % 26;
        if (divResult == 0) {
            return String.valueOf((char) (remainderResult + 65));
        } else {
            return transCharByNum(divResult - 1) + transCharByNum(remainderResult);
        }
    }

    private void setParentIsAppendPrefix(CascadeValidateModelWrapper cascadeValidateModel) {
        CascadeValidateItemWrapper parentValue = cascadeValidateModel.getParentValue();
        CascadeValidateItemWrapper lastValue = parentValue;
        while (parentValue != null) {
            lastValue = parentValue;
            parentValue = parentValue.getOwnModel().getParentValue();
        }
        lastValue.setAppendPrefix(true);
    }

    private String findParentValuePath(CascadeValidateModelWrapper cascadeValidateModel) {
        Stack<String> stack = new Stack<>();
        CascadeValidateItemWrapper parentValue = cascadeValidateModel.getParentValue();
        while (parentValue != null) {
            String value = parentValue.getValue();
            stack.push(value);
            if (parentValue.isAppendPrefix()) {
                stack.push("_");
            }
            parentValue = parentValue.getOwnModel().getParentValue();
        }
        if (!stack.isEmpty()) {
            StringBuffer stringBuffer = new StringBuffer();
            while (!stack.isEmpty()) {
                stringBuffer.append(stack.pop());
            }
            return stringBuffer.toString();
        }
        return null;
    }

    private DataValidationHelper createDataValidationHelper() {
        if (currentExcelType.equals(XLSX)) {
            if (isBigData) {
                return ((SXSSFSheet) sheet).getDataValidationHelper();
            } else {
                return new XSSFDataValidationHelper((XSSFSheet) sheet);
            }
        }
        return new HSSFDataValidationHelper((HSSFSheet) sheet);
    }
}
