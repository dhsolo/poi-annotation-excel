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

import io.github.dhsolo.poi.excel.model.ComplexExcelModel;
import io.github.dhsolo.poi.excel.ExcelModel;
import io.github.dhsolo.poi.excel.model.ExcelRowData;
import io.github.dhsolo.poi.excel.cascade.CascadeValidateModel;
import io.github.dhsolo.poi.excel.cascade.CascadeValidateModelWrapper;
import io.github.dhsolo.poi.excel.cascade.CascadeValidateItemWrapper;
import io.github.dhsolo.poi.excel.core.ExcelCascadeAble;
import io.github.dhsolo.poi.excel.exception.ExcelAnnotationException;
import io.github.dhsolo.poi.excel.exception.ExcelPropertyException;
import io.github.dhsolo.poi.excel.exception.ExcelReturnTypeException;
import io.github.dhsolo.poi.excel.validation.ExcelCustomValidate;
import io.github.dhsolo.common.Reflect;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * Default implementation of {@link AnnotationProcessor} that processes
 * Excel annotations on model classes via Java reflection.
 *
 * @author dhsolo
 * @since 1.0
 */
public class DefaultAnnotationProcessor implements AnnotationProcessor {

    private static final Logger logger = LoggerFactory.getLogger(DefaultAnnotationProcessor.class);
    private static final String LISTBOX_RETURN_TYPE_ERR =
            "Method '%s' annotated with @ExcelListBox must return List<String>, Set<String>, String[], or CascadeValidateModel";

    private final Object excelInfo;

    // Typed annotation metadata, populated by handleExcelInfoAnnotation().
    private final List<Field> columnFields = new ArrayList<>();
    private final List<Field> imageFields = new ArrayList<>();
    private final List<Field> dateFields = new ArrayList<>();
    private final List<Field> listBoxFields = new ArrayList<>();
    private final List<Field> infoChildFields = new ArrayList<>();
    private final List<Field> parentColumnFields = new ArrayList<>();
    private final List<Field> rowFields = new ArrayList<>();
    private final List<Method> translateMethods = new ArrayList<>();
    private final List<Method> listBoxMethods = new ArrayList<>();
    private final List<Method> customValidateMethods = new ArrayList<>();
    private Field titleField;
    private Field excelDataField;

    private ExcelAnnotationPropertyInfo info;

    /**
     * Creates a new processor for the given model instance and eagerly processes all
     * annotation metadata.
     *
     * @param excelInfo the model object whose class is annotated with {@link ExcelInfo};
     *                  must not be {@code null}
     * @throws ExcelAnnotationException if the class
     *         is missing the required {@link ExcelInfo} annotation
     */
    public DefaultAnnotationProcessor(Object excelInfo){
        this.excelInfo = excelInfo;
        handleExcelInfoAnnotation();
        handleExcelInfo();
    }


    /**
     * {@inheritDoc}
     *
     * <p>Returns the {@link ExcelAnnotationProperty} that was fully built during
     * construction. Subsequent calls return the same cached instance.
     */
    @Override
    public ExcelAnnotationProperty getExcelAnnotationProperty(){
        return info;
    }




    /**
     * Process all annotation metadata from the model class hierarchy.
     */
    private void handleExcelInfoAnnotation(){
        Class<?> clazz = excelInfo.getClass();
        while(clazz != null && clazz != Object.class){
            for (Field field : clazz.getDeclaredFields()) {
                if (field.getAnnotation(ExcelColumn.class) != null)   columnFields.add(field);
                if (field.getAnnotation(ExcelImage.class) != null)    imageFields.add(field);
                if (field.getAnnotation(ExcelDateFormat.class) != null) dateFields.add(field);
                if (field.getAnnotation(ExcelInfoChild.class) != null) infoChildFields.add(field);
                if (field.getAnnotation(ExcelColumnParent.class) != null) parentColumnFields.add(field);
                if (field.getAnnotation(ExcelListBox.class) != null)  listBoxFields.add(field);
                if (field.getAnnotation(ExcelRow.class) != null)      rowFields.add(field);
                if (field.getAnnotation(ExcelTitle.class) != null && titleField == null) {
                    titleField = field;
                }
                if (field.getAnnotation(ExcelData.class) != null && excelDataField == null) {
                    excelDataField = field;
                }
            }
            for (Method method : clazz.getDeclaredMethods()) {
                if (method.getAnnotation(ExcelTranslateMethod.class) != null)      translateMethods.add(method);
                if (method.getAnnotation(ExcelListBox.class) != null)              listBoxMethods.add(method);
                if (method.getAnnotation(ExcelCustomValidateMethod.class) != null) customValidateMethods.add(method);
            }
            clazz = clazz.getSuperclass();
        }

        info = new ExcelAnnotationPropertyInfo();
        if (excelInfo instanceof ComplexExcelModel) {
            List<?> complexModels = ((ComplexExcelModel) excelInfo).getComplexModels();
            if (complexModels != null && !complexModels.isEmpty()) {
                info.complexHandles = complexModels.stream()
                        .map(DefaultAnnotationProcessor::new)
                        .collect(Collectors.toList());
            }
        }
    }

    /**
     * Validate and bind the @ExcelInfo annotation.
     */
    private void handleExcelInfo(){
        ExcelInfo excelInfoAnno = this.excelInfo.getClass().getAnnotation(ExcelInfo.class);
        if (excelInfoAnno == null) {
            throw new ExcelAnnotationException(
                    "Missing @ExcelInfo annotation on class: " + this.excelInfo.getClass().getSimpleName());
        }
        info.excelInfo = excelInfoAnno;
        handleExcelTitle();
        handleExcelRows();
    }

    /**
     * Checks whether the model implements {@link ExcelCascadeAble}
     * and, if so, invokes {@code cascadeList()} to populate the cascade-validation model list.
     */
    @SuppressWarnings("unchecked")
    private void handleCascadeAbleInterface() {
        if (!ExcelCascadeAble.class.isAssignableFrom(excelInfo.getClass())) return;
        try {
            Method cascadeList = Reflect.findMethod(excelInfo.getClass(), "cascadeList");
            info.cascadeValidateModel = (List<CascadeValidateModel>) cascadeList.invoke(excelInfo);
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }


    /**
     * Process title annotation field.
     */
    private void handleExcelTitle(){
        if (titleField != null) {
            if (!String.class.isAssignableFrom(titleField.getType())) {
                throw new ExcelReturnTypeException("Fields annotated with @ExcelTitle must be of type String");
            }
            titleField.setAccessible(true);
            Object fieldValue = Reflect.getField(titleField, excelInfo);
            if (fieldValue != null) {
                info.title = (String) fieldValue;
            }
        }
        handleExcelColumn();
    }

    /**
     * Processes {@code @ExcelRow} annotated fields and builds the custom row configuration list.
     * Rows whose field value is null or empty are skipped and do not produce empty rows.
     */
    private void handleExcelRows() {
        if (rowFields.isEmpty()) return;

        // When there are multiple @ExcelRow fields, order values must all be distinct;
        // otherwise ordering depends on JVM implementation behavior.
        if (rowFields.size() > 1) {
            long distinctOrders = rowFields.stream()
                    .map(f -> f.getAnnotation(ExcelRow.class).order())
                    .distinct()
                    .count();
            if (distinctOrders != rowFields.size()) {
                throw new ExcelAnnotationException(
                        "Class " + excelInfo.getClass().getSimpleName() +
                        " has multiple @ExcelRow fields with duplicate 'order' values. " +
                        "Each @ExcelRow field must have a unique 'order' value to guarantee row ordering.");
            }
        }

        rowFields.sort(Comparator.comparingInt(f -> f.getAnnotation(ExcelRow.class).order()));

        List<ExcelAnnotationProperty.DiyRowConfig> configs = new ArrayList<>();
        for (Field field : rowFields) {
            field.setAccessible(true);
            ExcelCell[] cells = field.getAnnotation(ExcelRow.class).cells();
            String text = resolveRowText(field, cells);
            if (Reflect.hasText(text)) {
                configs.add(new ExcelAnnotationProperty.DiyRowConfig(text, /* merge= */ true));
            }
        }
        info.diyRows = configs;
    }

    private String resolveRowText(Field field, ExcelCell[] cells) {
        if (cells.length == 0) {
            Object val = Reflect.getField(field, excelInfo);
            return val != null ? val.toString() : null;
        }
        ExcelCell cellDef = cells[0];
        if (!cellDef.value().isEmpty()) {
            return cellDef.value();
        }
        if (!cellDef.field().isEmpty()) {
            Field targetField = Reflect.findField(excelInfo.getClass(), cellDef.field());
            if (targetField != null) {
                targetField.setAccessible(true);
                Object val = Reflect.getField(targetField, excelInfo);
                return val != null ? val.toString() : null;
            }
        }
        return null;
    }

    /**
     * Process column annotations and build column model list.
     */
    /** One column to render. */
    private static final class ColumnEntry {
        /** The declaring field; {@code null} for an {@code @ExcelColumnParent} child column. */
        final Field field;
        final ExcelColumn annotation;
        /** Value/identity key: the field name, or the sourceField for a parent child. */
        final String readKey;
        /** {@code @ExcelInfoChild} flatten prefix; {@code null} otherwise. */
        final String pathPrefix;
        /** {@code @ExcelColumnParent} group header label; {@code null} if not grouped. */
        final String parentLabel;
        /** Identity of the parent group (parent field name); {@code null} if not grouped. */
        final String groupKey;

        ColumnEntry(Field field, ExcelColumn annotation, String readKey,
                    String pathPrefix, String parentLabel, String groupKey) {
            this.field = field;
            this.annotation = annotation;
            this.readKey = readKey;
            this.pathPrefix = pathPrefix;
            this.parentLabel = parentLabel;
            this.groupKey = groupKey;
        }
    }

    /**
     * Collects all columns ordered by {@link ExcelColumn#index()}: top-level {@code @ExcelColumn}
     * fields, the columns of every {@code @ExcelInfoChild} nested object (flattened with a path
     * prefix), and the columns of every {@code @ExcelColumnParent} group (fieldless; value read
     * via {@code sourceField}/{@code sourcePath}).
     */
    private List<ColumnEntry> collectColumnEntries() {
        List<ColumnEntry> entries = new ArrayList<>();
        for (Field f : columnFields) {
            entries.add(new ColumnEntry(f, f.getAnnotation(ExcelColumn.class), f.getName(), null, null, null));
        }
        for (Field childField : infoChildFields) {
            collectChildColumns(childField.getType(), childField.getName(), entries, new HashSet<>());
        }
        for (Field parentField : parentColumnFields) {
            ExcelColumnParent parent = parentField.getAnnotation(ExcelColumnParent.class);
            String label = Reflect.hasText(parent.value()) ? parent.value() : parentField.getName();
            for (ExcelColumn ann : parent.columns()) {
                String readKey = Reflect.hasText(ann.sourceField()) ? ann.sourceField()
                        : (Reflect.hasText(ann.sourcePath()) ? ann.sourcePath() : ann.columnName());
                entries.add(new ColumnEntry(null, ann, readKey, null, label, parentField.getName()));
            }
        }
        entries.sort(Comparator.comparingInt(e -> e.annotation.index()));
        return entries;
    }

    /**
     * Recursively flattens an {@code @ExcelInfoChild} nested type: emits a column entry for each of
     * its {@code @ExcelColumn} fields (with the accumulated dot-path prefix) and descends into any
     * further {@code @ExcelInfoChild} fields. The {@code visited} set guards against type cycles.
     */
    private void collectChildColumns(Class<?> type, String prefix, List<ColumnEntry> entries, Set<Class<?>> visited) {
        if (type == null || !visited.add(type)) return;
        for (Class<?> c = type; c != null && c != Object.class; c = c.getSuperclass()) {
            for (Field cf : c.getDeclaredFields()) {
                ExcelColumn ann = cf.getAnnotation(ExcelColumn.class);
                if (ann != null) {
                    entries.add(new ColumnEntry(cf, ann, cf.getName(), prefix, null, null));
                }
                if (cf.getAnnotation(ExcelInfoChild.class) != null) {
                    collectChildColumns(cf.getType(), prefix + "." + cf.getName(), entries, visited);
                }
            }
        }
        visited.remove(type);
    }

    private void handleExcelColumn(){
        handleCascadeAbleInterface();
        if (columnFields.isEmpty() && infoChildFields.isEmpty() && parentColumnFields.isEmpty()) return;

        Map<String, CascadeValidateModel> cascadeValidateModelMap = info.cascadeValidateModel != null
                ? info.cascadeValidateModel.stream().collect(Collectors.toMap(CascadeValidateModel::getFieldName, Function.identity()))
                : new HashMap<>();

        // Build the full column list: top-level @ExcelColumn fields plus columns flattened from
        // @ExcelInfoChild nested objects, then order everything by @ExcelColumn.index().
        List<ColumnEntry> entries = collectColumnEntries();

        Map<String, ExcelModel> excelModelMap = new HashMap<>();
        int totalCells = 0;
        for (ColumnEntry e : entries) {
            if (e.field != null && e.pathPrefix == null) {
                excelModelMap.put(e.field.getName(), new ExcelModel(e.field.getName()));
            }
            int mci = e.annotation.mergeCellIndex();
            totalCells += (mci <= 0) ? 1 : mci;
        }

        String[] header = new String[totalCells];
        Map<Integer, Integer> columnWidth = new HashMap<>();
        Map<Integer, String> columnMergeInfo = new HashMap<>();
        LinkedList<ExcelModel> excelModels = new LinkedList<>();
        AtomicInteger count = new AtomicInteger();

        List<ExcelAnnotationProperty.ParentHeader> parentHeaders = new ArrayList<>();
        String curGroup = null, curLabel = null;
        int spanStart = -1, spanEnd = -1;
        Set<String> closedGroups = new HashSet<>();

        for (ColumnEntry entry : entries) {
            Field column = entry.field;
            ExcelColumn annotation = entry.annotation;
            String readKey = entry.readKey;
            boolean isInfoChild = entry.pathPrefix != null;
            boolean isParentChild = column == null;
            ExcelModel excelModel = (!isInfoChild && !isParentChild)
                    ? excelModelMap.get(readKey) : new ExcelModel(readKey);
            String columnName = annotation.columnName();

            int startCol = count.get();
            columnWidth.put(count.get(), annotation.columnWidth());
            if (annotation.needMergeCell()) {
                columnMergeInfo.put(count.get(), readKey);
            }
            excelModel.setNullAble(annotation.nullable());
            excelModel.setNoneCellDefaultValue(annotation.noneCellDefaultValue());
            excelModel.setMergeCellIndex(annotation.mergeCellIndex());

            if (isParentChild) {
                // @ExcelColumnParent child: value comes from sourceField/sourcePath on the row.
                applyInlineTranslate(annotation, excelModel);
                applySourceMapping(annotation, excelModel);
            } else if (isInfoChild) {
                excelModel.setSourcePath(entry.pathPrefix + "." + readKey);
                excelModel.setFlattened(true);
                applyInlineTranslate(annotation, excelModel);
                applyExcelImage(column, excelModel);
                applyExcelDateFormat(column, excelModel);
                applyExcelFormula(column, excelModel);
            } else {
                column.setAccessible(true);
                CascadeValidateModel cascadeModel = cascadeValidateModelMap.get(readKey);
                if (cascadeModel != null) {
                    excelModel.setCascadeValidateModel(createCascadeModelWrapper(excelModel, excelModelMap, cascadeModel));
                }

                applyInlineTranslate(annotation, excelModel);
                applySourceMapping(annotation, excelModel);
                applyExcelImage(column, excelModel);
                applyExcelListBoxField(column, readKey, excelModel);
                applyExcelListBoxMethod(readKey, excelModel, excelModelMap);
                applyExcelDateFormat(column, excelModel);
                applyExcelTranslateMethod(readKey, excelModel);
                applyExcelCustomValidateMethod(readKey, excelModel);
                applyExcelFormula(column, excelModel);

                if (column.getAnnotation(ExcelData.class) != null) {
                    info.excelData = Reflect.getField(column, excelInfo);
                }
            }

            excelModel.setIndex(annotation.index());
            excelModel.setColumnName(columnName);

            if (excelModel.getMergeCellIndex() > 1) {
                int mci = excelModel.getMergeCellIndex();
                boolean isFirst = true;
                for (int i = 0; i < mci; i++) {
                    header[count.getAndIncrement()] = columnName;
                    ExcelModel clone = Reflect.clone(excelModel);
                    if (!isFirst) clone.setMergeIndexEnd(true);
                    isFirst = false;
                    excelModels.add(clone);
                }
            } else {
                header[count.getAndIncrement()] = columnName;
                excelModels.add(excelModel);
            }

            // Track contiguous @ExcelColumnParent spans (in data-column space).
            int endCol = count.get() - 1;
            if (entry.groupKey != null) {
                if (entry.groupKey.equals(curGroup)) {
                    spanEnd = endCol;
                } else {
                    if (curGroup != null) {
                        parentHeaders.add(new ExcelAnnotationProperty.ParentHeader(curLabel, spanStart, spanEnd));
                        closedGroups.add(curGroup);
                    }
                    if (closedGroups.contains(entry.groupKey)) {
                        throw new ExcelAnnotationException(
                                "@ExcelColumnParent group '" + entry.parentLabel + "' has non-contiguous columns; "
                                        + "its child columns must use consecutive @ExcelColumn index values");
                    }
                    curGroup = entry.groupKey;
                    curLabel = entry.parentLabel;
                    spanStart = startCol;
                    spanEnd = endCol;
                }
            } else if (curGroup != null) {
                parentHeaders.add(new ExcelAnnotationProperty.ParentHeader(curLabel, spanStart, spanEnd));
                closedGroups.add(curGroup);
                curGroup = null;
            }
        }
        if (curGroup != null) parentHeaders.add(new ExcelAnnotationProperty.ParentHeader(curLabel, spanStart, spanEnd));

        info.header = header;
        info.excelModels = excelModels;
        info.mergeInfo = columnMergeInfo;
        info.columnWidthInfo = columnWidth;
        info.parentHeaders = parentHeaders.isEmpty() ? null : parentHeaders;
        if (excelDataField != null) {
            excelDataField.setAccessible(true);
            info.excelData = Reflect.getField(excelDataField, this.excelInfo);
        }
    }

    private void applyInlineTranslate(ExcelColumn annotation, ExcelModel excelModel) {
        String[] translateDefs = annotation.translate();
        if (translateDefs.length == 0) return;
        Map<Object, Object> translateMap = new LinkedHashMap<>();
        for (String def : translateDefs) {
            String[] kv = def.split(":", 2);
            if (kv.length == 2) {
                translateMap.put(kv[0].trim(), kv[1].trim());
            }
        }
        excelModel.setNeedtranslate(true);
        excelModel.setTranslateMappingInfo(translateMap);
    }

    private void applySourceMapping(ExcelColumn annotation, ExcelModel excelModel) {
        String sp = annotation.sourcePath();
        if (!sp.isEmpty()) excelModel.setSourcePath(sp);
        String sf = annotation.sourceField();
        if (!sf.isEmpty()) excelModel.setSourceField(sf);
    }

    private void applyExcelImage(Field column, ExcelModel excelModel) {
        ExcelImage image = column.getAnnotation(ExcelImage.class);
        if (image == null) return;
        excelModel.setPicture(true);
        excelModel.setImageDownPath(image.imageDownPath());
        excelModel.setImageVisitPrex(image.imageVisitPrev());
    }

    private void applyExcelListBoxField(Field column, String fieldName, ExcelModel excelModel) {
        if (listBoxFields.isEmpty()) return;
        // Match by field name (the @ExcelListBox annotation lives on the same field as @ExcelColumn)
        ExcelListBox excelListBox = column.getAnnotation(ExcelListBox.class);
        if (excelListBox == null) return;
        String[] listArray = excelListBox.listTextBox();
        if (listArray == null || listArray.length == 0) return;
        Set<String> listSet = Arrays.stream(listArray).collect(Collectors.toSet());
        excelModel.setCascadeValidateModel(
                createCascadeModel(listSet, excelModel, excelListBox.isNeedAddTranslationException()));
        excelModel.setListBox(true);
    }

    private void applyExcelListBoxMethod(String fieldName, ExcelModel excelModel, Map<String, ExcelModel> excelModelMap) {
        if (listBoxMethods.isEmpty()) return;

        List<String> blankColumnNameMethods = listBoxMethods.stream()
                .filter(m -> !Reflect.hasText(m.getAnnotation(ExcelListBox.class).columnName()))
                .map(Method::getName)
                .collect(Collectors.toList());
        if (!blankColumnNameMethods.isEmpty()) {
            throw new ExcelPropertyException(
                    "Methods " + blankColumnNameMethods + " annotated with @ExcelListBox must specify a non-blank columnName");
        }

        List<Method> matching = listBoxMethods.stream()
                .filter(m -> m.getAnnotation(ExcelListBox.class).columnName().equals(fieldName))
                .collect(Collectors.toList());
        if (matching.size() == 1) {
            invokeListBoxMethod(matching.get(0), excelModel, fieldName, excelModelMap);
        }
    }

    /**
     * Invokes a {@code @ExcelListBox}-annotated method and wires the returned values
     * into the given {@link ExcelModel}. Accepts {@code List<String>}, {@code Set<String>},
     * {@code String[]}, or {@link CascadeValidateModel} return types.
     */
    @SuppressWarnings("unchecked")
    private void invokeListBoxMethod(Method method, ExcelModel excelModel, String fieldName,
                                     Map<String, ExcelModel> excelModelMap) {
        Class<?> returnType = method.getReturnType();
        boolean isCascadeType = CascadeValidateModel.class.isAssignableFrom(returnType);
        try {
            if (CascadeValidateModel.class.isAssignableFrom(returnType)) {
                CascadeValidateModel cvm = (CascadeValidateModel) method.invoke(excelInfo);
                if (cvm != null) {
                    cvm.setFieldName(fieldName);
                    excelModel.setCascadeValidateModel(createCascadeModelWrapper(excelModel, excelModelMap, cvm));
                    excelModel.setListBox(true);
                }
                return;
            }

            Set<String> values = null;
            if (List.class.isAssignableFrom(returnType)) {
                List<String> list = (List<String>) method.invoke(excelInfo);
                if (list != null && !list.isEmpty()) values = new LinkedHashSet<>(list);
            } else if (Set.class.isAssignableFrom(returnType)) {
                values = (Set<String>) method.invoke(excelInfo);
            } else if (returnType.getComponentType() != null) {
                String[] arr = (String[]) method.invoke(excelInfo);
                if (arr != null && arr.length > 0) values = Arrays.stream(arr).collect(Collectors.toSet());
            } else {
                throw new ExcelReturnTypeException(String.format(LISTBOX_RETURN_TYPE_ERR, method.getName()));
            }
            if (values != null && !values.isEmpty()) {
                excelModel.setCascadeValidateModel(createCascadeModel(values, excelModel, true));
                excelModel.setListBox(true);
            }
        } catch (IllegalAccessException | InvocationTargetException e) {
            if (isCascadeType) throw new RuntimeException(e);
            logger.error("Failed to invoke @ExcelListBox method '{}'", method.getName(), e);
        } catch (ClassCastException e) {
            throw new ExcelReturnTypeException(String.format(LISTBOX_RETURN_TYPE_ERR, method.getName()));
        }
    }

    private void applyExcelDateFormat(Field column, ExcelModel excelModel) {
        ExcelDateFormat excelDateFormat = column.getAnnotation(ExcelDateFormat.class);
        if (excelDateFormat == null) return;
        excelModel.setDate(true);
        excelModel.setPattern(excelDateFormat.pattern());
    }

    @SuppressWarnings("unchecked")
    private void applyExcelCustomValidateMethod(String fieldName, ExcelModel excelModel) {
        if (customValidateMethods.isEmpty()) return;
        for (Method method : customValidateMethods) {
            if (!method.getAnnotation(ExcelCustomValidateMethod.class).columnName().equals(fieldName)) continue;
            if (!ExcelCustomValidate.class.isAssignableFrom(method.getReturnType())) {
                throw new ExcelReturnTypeException(
                        "Methods annotated with @ExcelCustomValidateMethod must return ExcelCustomValidate");
            }
            try {
                excelModel.setExcelCustomValidate((ExcelCustomValidate) Reflect.invokeMethod(method, excelInfo));
            } catch (ClassCastException e) {
                throw new ExcelReturnTypeException(
                        "Methods annotated with @ExcelCustomValidateMethod must return ExcelCustomValidate");
            }
            break;
        }
    }

    private void applyExcelTranslateMethod(String fieldName, ExcelModel excelModel) {
        if (translateMethods.isEmpty()) return;
        Method matched = null;
        for (Method method : translateMethods) {
            if (!method.getAnnotation(ExcelTranslateMethod.class).columnName().equals(fieldName)) continue;
            if (!Function.class.isAssignableFrom(method.getReturnType())) {
                throw new ExcelReturnTypeException("Methods annotated with @ExcelTranslateMethod must return Function<ExcelRowData, Object>");
            }
            try {
                excelModel.setBiFunction((Function<ExcelRowData, Object>) Reflect.invokeMethod(method, excelInfo));
            } catch (ClassCastException e) {
                throw new ExcelReturnTypeException("Methods annotated with @ExcelTranslateMethod must return Function<ExcelRowData, Object>");
            }
            matched = method;
            break;
        }
        if (matched != null) translateMethods.remove(matched);
    }

    private void applyExcelFormula(Field column, ExcelModel excelModel) {
        ExcelFormula excelFormula = column.getAnnotation(ExcelFormula.class);
        if (excelFormula != null) {
            excelModel.setStrFormula(excelFormula.formula());
        }
    }

    /**
     * Recursively wraps a {@link CascadeValidateModel} and its nested items into the
     * internal {@link CascadeValidateModelWrapper} structure used by the export/import engine.
     *
     * @param excelModel    the {@link ExcelModel} that owns this cascade constraint
     * @param excelModelMap map from field name to {@link ExcelModel} for the current sheet,
     *                      used to resolve child cascade references
     * @param cascadeModel  the cascade validation model to wrap
     * @return a fully populated {@link CascadeValidateModelWrapper}
     */
    private CascadeValidateModelWrapper createCascadeModelWrapper(ExcelModel excelModel, Map<String, ExcelModel> excelModelMap,CascadeValidateModel cascadeModel) {
        CascadeValidateModelWrapper cascadeValidateModelWrapper = new CascadeValidateModelWrapper();
        cascadeValidateModelWrapper.setExcelModel(excelModel);
        cascadeValidateModelWrapper.setNeedAddTranslationException(cascadeModel.isNeedAddTranslationException());
        List<CascadeValidateModel.CascadeValidateItem> items = cascadeModel.getItems();
        if(items!=null && !items.isEmpty()){
            List<CascadeValidateItemWrapper> cascadeValidateItemWrappers = new ArrayList<>(items.size());
            cascadeValidateModelWrapper.setItems(cascadeValidateItemWrappers);
            for(CascadeValidateModel.CascadeValidateItem item : items){
                CascadeValidateItemWrapper cascadeValidateItemWrapper = new CascadeValidateItemWrapper();
                if(item.getCascadeValidateModels()!=null && !item.getCascadeValidateModels().isEmpty()){
                    cascadeValidateModelWrapper.setHasCascade(true);
                    List<CascadeValidateModelWrapper> cascadeModelWrappers = new ArrayList<>(item.getCascadeValidateModels().size());
                    for(CascadeValidateModel itemChild : item.getCascadeValidateModels()){
                        CascadeValidateModelWrapper cascadeModelWrapper = createCascadeModelWrapper(excelModelMap.get(itemChild.getFieldName()), excelModelMap, itemChild);
                        cascadeModelWrapper.setParentValue(cascadeValidateItemWrapper);
                        cascadeModelWrappers.add(cascadeModelWrapper);
                    }
                    cascadeValidateItemWrapper.setCascadeValidateModels(cascadeModelWrappers);
                }
                cascadeValidateItemWrapper.setValue(item.getValue());
                cascadeValidateItemWrapper.setOwnModel(cascadeValidateModelWrapper);
                cascadeValidateItemWrappers.add(cascadeValidateItemWrapper);
            }
        }
        return cascadeValidateModelWrapper;
    }

    /**
     * Builds a flat {@link CascadeValidateModelWrapper} from a set of allowed string values,
     * typically sourced from a static or dynamic {@link ExcelListBox} declaration.
     *
     * @param collect1                  the set of allowed cell values
     * @param excelModel                the {@link ExcelModel} that this validation applies to
     * @param needAddTranslationException whether to raise an exception when an imported value
     *                                  is not in the allowed set
     * @return a {@link CascadeValidateModelWrapper} containing one item per allowed value
     */
    private CascadeValidateModelWrapper createCascadeModel(Set<String> collect1,ExcelModel excelModel,boolean needAddTranslationException) {
        CascadeValidateModelWrapper cascadeValidateModel = new CascadeValidateModelWrapper();
        cascadeValidateModel.setExcelModel(excelModel);
        cascadeValidateModel.setNeedAddTranslationException(needAddTranslationException);
        List<CascadeValidateItemWrapper> items = new ArrayList<>(collect1.size());
        collect1.forEach(item->{
            CascadeValidateItemWrapper cascadeValidateItem = new CascadeValidateItemWrapper();
            cascadeValidateItem.setValue(item);
            items.add(cascadeValidateItem);
        });
        cascadeValidateModel.setItems(items);
        return cascadeValidateModel;
    }

    /**
     * Concrete {@link ExcelAnnotationProperty} implementation populated by
     * {@link DefaultAnnotationProcessor} during annotation processing.
     *
     * <p>All fields are package-private and set directly by the enclosing processor;
     * the public getters expose the fully resolved configuration to the export/import engine.
     */
    public static class ExcelAnnotationPropertyInfo implements ExcelAnnotationProperty{

        private String title;

        private String [] header;

        private List<ExcelModel> excelModels;

        private ExcelInfo excelInfo;

        private Object excelData;

        private Map<Integer,String> mergeInfo;

        private Map<Integer,Integer> columnWidthInfo;

        private List<CascadeValidateModel> cascadeValidateModel;

        private List<AnnotationProcessor> complexHandles;

        private List<ExcelAnnotationProperty.DiyRowConfig> diyRows;

        private List<ExcelAnnotationProperty.ParentHeader> parentHeaders;

        /** Returns the title text to render above the header row, or {@code null} if absent. */
        @Override
        public String getTitle() {
            return title;
        }

        /** Returns the ordered array of column header labels, one entry per logical column. */
        @Override
        public String[] getHeader() {
            return header;
        }

        /** Returns the parent-header groups from {@code @ExcelColumnParent}, or {@code null}. */
        @Override
        public List<ExcelAnnotationProperty.ParentHeader> getParentHeaders() {
            return parentHeaders;
        }

        /** Returns the ordered list of {@link ExcelModel} descriptors, one per logical column. */
        @Override
        public List<ExcelModel> getExcelModels() {
            return excelModels;
        }

        /** Returns the {@link ExcelInfo} annotation instance from the model class. */
        @Override
        public ExcelInfo getExcelInfo() {
            return excelInfo;
        }

        /**
         * Returns the data collection (typically a {@code List}) read from the
         * {@link ExcelData}-annotated field, or {@code null} if the field is absent or empty.
         */
        @Override
        public Object getExcelData() {
            return excelData;
        }

        /**
         * Returns a map from zero-based column index to field name for columns that have
         * row-level cell merging enabled ({@code needMergeCell = true}).
         */
        @Override
        public Map<Integer, String> getMergeInfo() {
            return mergeInfo;
        }

        /**
         * Returns a map from zero-based column index to character width for columns that
         * declare a custom {@link ExcelColumn#columnWidth()}.
         */
        @Override
        public Map<Integer, Integer> getColumnWidthInfo() {
            return columnWidthInfo;
        }

        /**
         * Returns the cascade-validation model list populated when the model implements
         * {@link ExcelCascadeAble}, or {@code null} otherwise.
         */
        @Override
        public List<CascadeValidateModel> getCascadeValidateModel() {
            return cascadeValidateModel;
        }

        /**
         * Returns one {@link AnnotationProcessor} per sub-sheet for models that implement
         * {@link io.github.dhsolo.poi.excel.model.ComplexExcelModel}, or {@code null} for simple models.
         */
        @Override
        public List<AnnotationProcessor> getComplexExcelModels() {
            return complexHandles;
        }

        /**
         * Returns the list of custom row configurations derived from {@link ExcelRow}-annotated
         * fields, in ascending {@link ExcelRow#order()} order. Returns {@code null} or an
         * empty list when no {@code @ExcelRow} fields are present.
         */
        @Override
        public List<ExcelAnnotationProperty.DiyRowConfig> getDiyRows() {
            return diyRows;
        }
    }

}
