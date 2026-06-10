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

import io.github.dhsolo.poi.excel.cascade.CascadeValidateModelWrapper;
import io.github.dhsolo.poi.excel.model.ExcelRowData;
import io.github.dhsolo.poi.excel.render.ExcelTranslateHandler;
import io.github.dhsolo.poi.excel.validation.ExcelCustomValidate;
import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Function;

/**
 * Describes a column mapping for Excel generation, including field name,
 * translation rules, image handling, and date formatting.
 *
 * @author dhsolo
 * @since 1.0
 */
public class ExcelModel implements ExcelTranslateHandler, Serializable {

	/** Column index. */
	private int index;

	private int realIndex;

    /** Name of the field to read values from. */
    private String fieldName;

	/** Column display name. */
    private String columnName;

	/**
	 * Default cell value when no data is available.
	 */
	private String noneCellDefaultValue;

	private ExcelModel parent;

    /** Whether the value needs to be translated to a custom value. */
    private boolean needtranslate=false;

	private int mergeCellIndex = 1;

	/**
	 * Cascade validation object.
	 */
	private CascadeValidateModelWrapper cascadeValidateModel;

	/**
	 * Custom translation function.
	 */
	private Function<ExcelRowData,Object> biFunction=null;

    /** Mapping of values to translate. */
    private Map<Object,Object> translateMappingInfo = new HashMap<>();

    /** Whether this field represents an image. */
    private boolean isPicture=false;

    /** Whether this field is a date/time type. */
    private boolean isDate=false;

    /** Whether this column is flattened from an @ExcelInfoChild nested object (value read/written via sourcePath). */
    private boolean flattened=false;

    public boolean isFlattened() { return flattened; }

    public ExcelModel setFlattened(boolean flattened) { this.flattened = flattened; return this; }

    /** Date format pattern. */
    private String pattern="yyyy-MM-dd HH:mm:ss";

    /** Image URL prefix (IP and port information). */
    private String imageVisitPrex;

    /** Image download path; required when images need to be saved during Excel upload. */
    private String imageDownPath;

    /** Whether the field may be null; defaults to false (not nullable). */
    private boolean nullAble=false;

    private boolean needHandle = false;

    /** Whether to record an error message when value translation fails for this field. */
    private boolean needAddTranslationException = true;

    private boolean isInteger = false;

    private boolean isFloat = false;

    private boolean isDouble = false;

    private ExcelCustomValidate excelCustomValidate;

	private boolean isListBox = false;

	private String strFormula;

	private boolean isMergeIndexEnd = false;

	/** Dot-notation path for extracting a value from the row data object (e.g. {@code "pointInfo.measureType"}). */
	private String sourcePath;

	/** Same-level field redirection: reads another field of the row data object to populate this column (e.g. {@code "correctResult"}). */
	private String sourceField;


	public int getRealIndex() {
		return realIndex;
	}

	public void setRealIndex(int realIndex) {
		this.realIndex = realIndex;
	}

	public String getStrFormula() {
		return strFormula;
	}

	public ExcelModel setStrFormula(String strFormula) {
		this.strFormula = strFormula;
		return this;
	}


	public ExcelModel setListBox(boolean listBox) {
		isListBox = listBox;
		return this;
	}

	public boolean isListBox() {
		return isListBox;
	}


	public ExcelModel setInteger(boolean integer) {
		isInteger = integer;
		return this;
	}

	public ExcelModel setFloat(boolean aFloat) {
		isFloat = aFloat;
		return this;
	}

	public ExcelModel setDouble(boolean aDouble) {
		isDouble = aDouble;
		return this;
	}

	public ExcelModel setExcelCustomValidate(ExcelCustomValidate excelCustomValidate) {
		this.excelCustomValidate = excelCustomValidate;
		return this;
	}

	public ExcelCustomValidate getExcelCustomValidate() {
		return excelCustomValidate;
	}

	/**
	 * Sets a custom translation function.
	 */
	public  ExcelModel  setBiFunction(Function<ExcelRowData,Object> biFunction){
    	if(biFunction!=null){
			this.biFunction = biFunction;
			needHandle=true;
		}
		return this;
	}

    public ExcelModel(String fieldName, boolean needtranslate, Map<Object, Object> translateMappingInfo, boolean isPicture) {
        this.fieldName = fieldName;
        this.needtranslate = needtranslate;
        this.translateMappingInfo = translateMappingInfo;
        this.isPicture = isPicture;
    }

	@Override
	public Object handler(ExcelRowData rowData) {
		if(needHandle())
			return biFunction.apply(rowData);
		throw new UnsupportedOperationException();
	}

	@Override
	public boolean needHandle() {
		return needHandle;
	}

	public boolean isNullAble() {
		return nullAble;
	}


	public ExcelModel setNullAble(boolean nullAble) {
		this.nullAble = nullAble;
		return this;
	}





	/**
	 * Shorthand factory for a plain column definition; more concise than
	 * {@code ExcelImportor.generateExcelModel()} when used with {@link ExcelUtil#importExcel}.
	 *
	 * <pre>
	 * ExcelUtil.importExcel(in, User.class,
	 *     ExcelModel.of("name"), ExcelModel.of("age"), ExcelModel.of("dept"));
	 * </pre>
	 *
	 * @param fieldName Java field name
	 */
	public static ExcelModel of(String fieldName) {
		return new ExcelModel(fieldName);
	}

	/**
	 * Shorthand factory for a column definition with an explicit display header.
	 * Useful when building column mappings without an annotation model:
	 *
	 * <pre>
	 * ExcelCreatorBuilder.create("Sheet1")
	 *     .data(list)
	 *     .addColumn(0, ExcelModel.of("name", "Employee Name"))
	 *     .addColumn(1, ExcelModel.of("dept", "Department"));
	 * </pre>
	 *
	 * @param fieldName  Java field name
	 * @param columnName display column header
	 */
	public static ExcelModel of(String fieldName, String columnName) {
		ExcelModel m = new ExcelModel(fieldName);
		m.setColumnName(columnName);
		return m;
	}

	/**
	 * Shorthand factory for a column definition with a translation mapping.
	 *
	 * @param fieldName            Java field name
	 * @param translateMappingInfo value translation mapping (Excel value → Java value)
	 */
	public static ExcelModel of(String fieldName, Map<Object, Object> translateMappingInfo) {
		return new ExcelModel(fieldName, translateMappingInfo);
	}

	/**
     * Constructor for a plain field.
     */
    public ExcelModel(String fieldName) {
		this(fieldName,false);
	}

	/**
	 * Constructor for dynamically added columns.
	 * @param fieldName field name
	 * @param columnName column display name
	 * @param index actual index
	 */
	public ExcelModel(String fieldName,String columnName,Integer index) {
		this(fieldName,false);
		this.setColumnName(columnName);
		this.setIndex(index);
	}

	public ExcelModel(String fieldName, boolean nullAble){
		super();
		this.fieldName = fieldName;
		this.nullAble = nullAble;
	}
    
    
    
    
    /**
     * Constructor for fields that require value translation.
     * @param fieldName field name
     * @param translateMappingInfo custom translation map
     */
	public ExcelModel(String fieldName, Map<Object, Object> translateMappingInfo) {
		super();
		this.fieldName = fieldName;
		this.needtranslate=true;
		this.translateMappingInfo = translateMappingInfo;
	}
	
	


	/**
	 * Constructor for image fields.
	 * @param fieldName field name
	 * @param imageVisitPrex image access URL prefix
	 * @param imageDownPath image download path
	 */
	public ExcelModel(String fieldName, String imageVisitPrex, String imageDownPath) {
		super();
		this.fieldName = fieldName;
		this.imageVisitPrex = imageVisitPrex;
		this.imageDownPath = imageDownPath;
		this.isPicture=true;
	}

	
	public ExcelModel(String fieldName, String pattern) {
		super();
		this.fieldName = fieldName;
		this.isDate = true;
		this.pattern=pattern;
	}


	public ExcelModel(String fieldName, boolean needtranslate, Map<Object, Object> translateMappingInfo,
			boolean isPicture, String imageDownPath,String imageVisitPrex) {
		super();
		this.fieldName = fieldName;
		this.needtranslate = needtranslate;
		this.translateMappingInfo = translateMappingInfo;
		this.isPicture = isPicture;
		this.imageDownPath = imageDownPath;
		this.imageVisitPrex=imageVisitPrex;
	}



	public ExcelModel setImageDownPath(String imageDownPath) {
    	this.imageDownPath=imageDownPath;
		return this;
    }

    public String getFieldName() {
        return fieldName;
    }

    public ExcelModel setFieldName(String fieldName) {
        this.fieldName = fieldName;
		return this;
    }

    public boolean isNeedtranslate() {
        return needtranslate;
    }

    public ExcelModel setNeedtranslate(boolean needtranslate) {
        this.needtranslate = needtranslate;
		return this;
    }

    public Map<Object, Object> getTranslateMappingInfo() {
        return translateMappingInfo;
    }

    public ExcelModel setTranslateMappingInfo(Map<Object, Object> translateMappingInfo) {
        this.translateMappingInfo = translateMappingInfo;
		return this;
    }

    public boolean isPicture() {
        return isPicture;
    }

    public ExcelModel setPicture(boolean picture) {
        isPicture = picture;
		return this;
    }



	public String getImageDownPath() {
		return imageDownPath;
	}



	public String getImageVisitPrex() {
		return imageVisitPrex;
	}



	public ExcelModel setImageVisitPrex(String imageVisitPrex) {
		this.imageVisitPrex = imageVisitPrex;
		return this;
	}
    
	public ExcelModel setDate(boolean isDate) {
		this.isDate=isDate;
		return this;
	}
    
	public boolean isDate() {
		return isDate;
	}


	public String getPattern() {
		return pattern;
	}


	public ExcelModel setPattern(String pattern) {
		isDate=true;
		this.pattern = pattern;
		return this;
	}


	public boolean isNeedAddTranslationException() {
		return needAddTranslationException;
	}

	public ExcelModel setNeedAddTranslationException(boolean needAddTranslationException) {
		this.needAddTranslationException = needAddTranslationException;
		return this;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public String getColumnName() {
		return columnName;
	}

	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}


	public void setMergeCellIndex(int mergeCellIndex) {
		this.mergeCellIndex = mergeCellIndex;
	}

	public int getMergeCellIndex() {
		return mergeCellIndex;
	}

	public void setMergeIndexEnd(boolean mergeIndexEnd) {
		isMergeIndexEnd = mergeIndexEnd;
	}

	public  boolean getMergeIndexEnd(){
		return isMergeIndexEnd;
	}


	public CascadeValidateModelWrapper getCascadeValidateModel() {
		return cascadeValidateModel;
	}

	public void setCascadeValidateModel(CascadeValidateModelWrapper cascadeValidateModel) {
		if(cascadeValidateModel!= null){
			isListBox = true;
			this.cascadeValidateModel = cascadeValidateModel;
		}
	}

	public void setNoneCellDefaultValue(String noneCellDefaultValue) {
		this.noneCellDefaultValue = noneCellDefaultValue;
	}

	public String getNoneCellDefaultValue() {
		return noneCellDefaultValue;
	}

	public String getSourcePath() { return sourcePath; }
	public ExcelModel setSourcePath(String sourcePath) { this.sourcePath = sourcePath; return this; }

	public String getSourceField() { return sourceField; }
	public ExcelModel setSourceField(String sourceField) { this.sourceField = sourceField; return this; }

	/** Returns {@code true} when {@link #sourcePath} is configured. */
	public boolean hasSourcePath() { return sourcePath != null && !sourcePath.isEmpty(); }

	/** Returns the effective source field name: returns {@link #sourceField} if configured, otherwise {@link #fieldName}. */
	public String effectiveSourceField() { return sourceField != null && !sourceField.isEmpty() ? sourceField : fieldName; }

	/**
	 * Shallow copy used for the column-expansion clones of {@code mergeCellIndex > 1}.
	 * Reference members (translate map, handler function, cascade/custom-validate wrappers)
	 * are shared, not deep-cloned — every clone describes the same logical column. This
	 * replaces Java-serialisation cloning, which failed with NotSerializableException for
	 * non-serialisable members such as {@code @ExcelTranslateMethod} lambdas.
	 *
	 * @return a field-by-field shallow copy of this model
	 */
	public ExcelModel copy() {
		ExcelModel c = new ExcelModel(this.fieldName);
		c.index = this.index;
		c.realIndex = this.realIndex;
		c.columnName = this.columnName;
		c.noneCellDefaultValue = this.noneCellDefaultValue;
		c.parent = this.parent;
		c.needtranslate = this.needtranslate;
		c.mergeCellIndex = this.mergeCellIndex;
		c.cascadeValidateModel = this.cascadeValidateModel;
		c.biFunction = this.biFunction;
		c.translateMappingInfo = this.translateMappingInfo;
		c.isPicture = this.isPicture;
		c.isDate = this.isDate;
		c.flattened = this.flattened;
		c.pattern = this.pattern;
		c.imageVisitPrex = this.imageVisitPrex;
		c.imageDownPath = this.imageDownPath;
		c.nullAble = this.nullAble;
		c.needHandle = this.needHandle;
		c.needAddTranslationException = this.needAddTranslationException;
		c.isInteger = this.isInteger;
		c.isFloat = this.isFloat;
		c.isDouble = this.isDouble;
		c.excelCustomValidate = this.excelCustomValidate;
		c.isListBox = this.isListBox;
		c.strFormula = this.strFormula;
		c.isMergeIndexEnd = this.isMergeIndexEnd;
		c.sourcePath = this.sourcePath;
		c.sourceField = this.sourceField;
		return c;
	}
}
