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

import java.awt.image.BufferedImage;
import java.io.*;
import java.lang.reflect.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import javax.imageio.ImageIO;

import io.github.dhsolo.poi.excel.ExcelCustomModel;
import io.github.dhsolo.poi.excel.validation.ExcelCustomValidate;
import io.github.dhsolo.poi.excel.ExcelModel;
import io.github.dhsolo.poi.excel.model.ExcelRowData;
import io.github.dhsolo.poi.excel.core.CSVWorkbook;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import io.github.dhsolo.common.CommonUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import io.github.dhsolo.common.Reflect;


/**
 * Excel import utility class.
 *
 * @author dhsolo
 * @since 1.0
 */
@SuppressWarnings({ "rawtypes", "unchecked" })
public class ExcelImportor {


	private static final Logger logger = LoggerFactory.getLogger(ExcelImportor.class);


	/** File input stream. */
	private InputStream in;

	/** Used to detect the type of the incoming file. */
	private byte[] header = new byte[8];

	/** Current workbook instance. */
	private Workbook workbook;

	/** All sheet pages. */
	private LinkedList<Sheet> sheets = new LinkedList<>();

	/** Current sheet. */
	private Sheet sheet;

	/** Current row. */
	private Row row;

	/** Current cell. */
	private Cell cell;

	/** Row index from which data reading starts. */
	private int startRow = 1;

	/** Start row index for each sheet page. */
	private LinkedList<Integer> sheetStartRow = new LinkedList<>();

	/** Column configuration list. */
	private LinkedList<LinkedList<ExcelModel>> columnNameList = new LinkedList<>();

	/** Column indices to skip during parsing. */
	private Map<Integer, List<Integer>> exceptColumnNumMap = new HashMap<>();

	/** Parsed data results. */
	private LinkedList<LinkedList<Map<String, Object>>> datas = new LinkedList<>();

	/** Error messages collected during parsing. */
	private StringBuilder errorMessage = new StringBuilder();

	/** Optional row-by-row listener; null means results are collected in {@link #datas}. */
	private ExcelReadListener readListener;

	private FormulaEvaluator formulaEvaluator;


	public ExcelImportor(Object excelModel){

	}

	public ExcelImportor(InputStream in) {
		try {
			ByteArrayOutputStream outPut = cloneInputStream(in);
			if (outPut == null) {
				throw new IOException("Failed to clone input stream; cannot read Excel file");
			}
			ByteArrayInputStream byteIn = new ByteArrayInputStream(outPut.toByteArray());
			header = byteIn.readNBytes(8);
			this.in = new ByteArrayInputStream(outPut.toByteArray());
		} catch (IOException e) {
			logger.error("Failed to initialise ExcelImportor from input stream", e);
		}

		init();
	}

	private FileMagic detectMagic() {
		try {
			return FileMagic.valueOf(new ByteArrayInputStream(header));
		} catch (IOException e) {
			return FileMagic.UNKNOWN;
		}
	}

	private boolean isXLS() { return detectMagic() == FileMagic.OLE2; }
	private boolean isXLSX() { return detectMagic() == FileMagic.OOXML; }

	public String getWorkBookType(){
		FileMagic fm = detectMagic();
		if (fm == FileMagic.OLE2) return "xls";
		if (fm == FileMagic.OOXML) return "xlsx";
		return null;
	}

	/**
	 * Returns the error message describing parse failures.
	 */
	public String getErrorMessage() {
		return errorMessage.toString();
	}

	/**
	 * Clones a byte input stream.
	 */
	private ByteArrayOutputStream cloneInputStream(InputStream input) {
		try {
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			byte[] buffer = new byte[1024];
			int len;
			while ((len = input.read(buffer)) > -1) {
				baos.write(buffer, 0, len);
			}
			baos.flush();
			return baos;
		} catch (IOException e) {
			logger.error("Failed to clone input stream", e);
			return null;
		}
	}

	/**
	 * Initializes the workbook from the input stream.
	 */
	private void init() {

		try {
			workbook = createWorkbook();
			if (isXLS()) {
				formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
			} else if (isXLSX()) {
				formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
			}
			int sheetNum = workbook.getNumberOfSheets();
			for (int i = 0; i < sheetNum; i++) {
				Sheet sheetAt = workbook.getSheetAt(i);
				if(!sheetAt.getSheetName().equals("listConstantData")){
					sheets.add(sheetAt);
				}
			}
		} catch (IOException e) {
			logger.error("Failed to create workbook from input stream", e);
		}
	}


	private Workbook createWorkbook() throws IOException {
		if (isXLS() || isXLSX()) {
			return WorkbookFactory.create(this.in);
		} else {
			return createCSVWorkbook(this.in);
		}
	}

	private Workbook createCSVWorkbook(InputStream in) throws IOException {
		return 	new CSVWorkbook(in);
	}


	public String getSheetName(int i){
		return workbook.getSheetName(i);
	}



	/**
	 * Parses the Excel file.
	 */
	public boolean analysisExcel() {
		return analysisExcel(true);
	}

	/**
	 * Parses the Excel file.
	 * When {@code firstErrorBreak} is {@code true}, parsing stops on the first error;
	 * when {@code false}, all data is parsed before returning.
	 */
	public boolean analysisExcel(boolean firstErrorBreak) {
		int listenerRowCount = 0;
		// Iterate over each sheet and extract its data
		for (int s = 0; s < sheets.size(); s++) {
			List<Integer> exceptColumnNum = exceptColumnNumMap.get(s);
			sheet = sheets.get(s);
			if (s < sheetStartRow.size()) {
				startRow = sheetStartRow.get(s);
			}
			int realRows = sheet.getLastRowNum();
			if(startRow > realRows){
				datas.add(new LinkedList<>());
				continue;
			}
			startRow = Math.min(startRow, sheet.getLastRowNum());
			int lastRowNum = sheet.getLastRowNum();
			LinkedList<Map<String, Object>> rowData = new LinkedList<>();
			for (int r = startRow; r <= lastRowNum; r++) {
				row = sheet.getRow(r);
				if (row == null ||  checkIsEmpty(row)) {
					continue;
				}
				// Per-row snapshots for @ExcelCustomValidateMethod: rebuilt every row so the
				// validator always sees the current row's cell/value, not the first row's.
				List<ExcelCustomModel> excelCustomModels = new ArrayList<>();
				short minColIx = row.getFirstCellNum();
				short maxColIx = row.getLastCellNum();
				if(minColIx <0 || maxColIx < 0){
					continue;
				}
				if(minColIx!=0 && (exceptColumnNum == null || !exceptColumnNum.contains(0))){
					minColIx = 0;
				}
				if(exceptColumnNum!=null && maxColIx<exceptColumnNum.size()) {
					errorMessage.append("Column count in Excel does not match the configured column definitions");
					return false;
				}
				int physicalIndex = 0;
				int errorLengthBeforeRow = errorMessage.length();
				Map<String, Object> cellData = new LinkedHashMap<>();
				Map<String, Object> originalCellData = new LinkedHashMap<>();
				for (short colIx = minColIx; colIx <= maxColIx; colIx++) {
					cell = row.getCell(colIx);
					if (s >= columnNameList.size())
						continue;

					if (physicalIndex > columnNameList.get(s).size() - 1)
						continue;
//					physicalIndex = Math.min(columnNameList.get(s).size()-1, physicalIndex);
					ExcelModel excelModel = columnNameList.get(s).get(physicalIndex);
					String fieldName = excelModel.getFieldName();
					// Flattened (@ExcelInfoChild) columns key by their distinct sourcePath so children of
					// different parents sharing a child field name do not collide.
					String dataKey = (excelModel.isFlattened() && excelModel.getSourcePath() != null)
							? excelModel.getSourcePath() : fieldName;
					Object value = null;

					if (exceptColumnNum!=null && exceptColumnNum.contains((int) colIx)) {
						continue;
					}
					if (cell == null && !excelModel.isPicture()) {
						if(!excelModel.isNullAble()){
							String errMsg = "Sheet '" + sheet.getSheetName() + "': row " + (r + 1) + ", column " + (colIx + 1) + " must not be empty";
							errorMessage.append(errMsg).append("\n");
							if (readListener != null) readListener.onError(errMsg, r - startRow);
							if(firstErrorBreak){
								return false;
							}
						}
						physicalIndex++;
						continue;
					}
					try {
						if (cell != null)
							value = getCellValue(cell);

						if(!excelModel.isNullAble() && (value ==null || (value+"").trim().length()==0)) {
							String errMsg = "Sheet '" + sheet.getSheetName() + "': row " + (r + 1) + ", column " + (colIx + 1) + " must not be empty";
							errorMessage.append(errMsg).append("\n");
							if (readListener != null) readListener.onError(errMsg, r - startRow);
							if(firstErrorBreak){
								return false;
							}
						}
					} catch (Exception e) {
						logger.error("Failed to read cell value", e);
						String errMsg = "Sheet '" + sheet.getSheetName() + "': row " + (r + 1) + ", column " + (colIx + 1) + " — failed to read cell value. Please verify the file format matches the template. Cause: " + e.getMessage();
						errorMessage.append(errMsg).append("\n");
						if (readListener != null) readListener.onError(errMsg, r - startRow);
						if(firstErrorBreak){
							return false;
						}
					}
					try {
						if (excelModel.isDate() && value != null) {
							value = DateUtil.getJavaDate(Double.parseDouble(value.toString()) );
							value = CommonUtil.formatDate(value, excelModel.getPattern());
						}
					} catch (Exception e) {
						logger.error("Date conversion failed", e);
					}
					originalCellData.put(dataKey, value);
					if (value!=null && value.toString().trim().length()>0 && excelModel.isNeedtranslate() && excelModel.getTranslateMappingInfo() != null) {
						Object temp = getFromMap(excelModel.getTranslateMappingInfo(), value, null);
						final Object originalValue = value;
						if (temp != null) {
							value = temp;
						}else if(excelModel.isNeedAddTranslationException()){
							String errMsg = "Sheet '" + sheet.getSheetName() + "': row " + (r + 1) + ", column " + (colIx + 1) + " — value translation failed. Allowed values are: " + excelModel.getTranslateMappingInfo().keySet();
							errorMessage.append(errMsg).append("\n");
							if (readListener != null) readListener.onError(errMsg, r - startRow);
							if(firstErrorBreak){
								return false;
							}
						}
					} else if (excelModel.isPicture()) {
						String downPath = excelModel.getImageDownPath();
						if (downPath == null) {
							throw new RuntimeException("Image download path is not configured on ExcelModel");
						}
						value = getPicture(r, colIx, downPath, excelModel.getImageVisitPrex());
					}
					if(excelModel.getExcelCustomValidate()!=null){
						ExcelCustomModel excelCustomModel = new ExcelCustomModel();
						excelCustomModel.setExcelModel(excelModel);
						excelCustomModel.setCurrentCellNum(colIx);
						excelCustomModel.setCurrentRowNum(r);
						excelCustomModel.setCell(cell);
						excelCustomModel.setRow(row);
						excelCustomModel.setCurrentValue(value);
						excelCustomModels.add(excelCustomModel);
					}

					cellData.put(dataKey, value);
					physicalIndex++;
				}
				Map<String,Object> finalCellData = new HashMap<>();
				Map<String,Object> finalOriginalCellData = new HashMap<>();
				finalCellData.putAll(cellData);
				finalOriginalCellData.putAll(originalCellData);
				if(excelCustomModels.size() > 0  && errorMessage.length() == errorLengthBeforeRow){
					for (ExcelCustomModel excelCustomModel : excelCustomModels){
						if(errorMessage.length() > errorLengthBeforeRow){
							if(firstErrorBreak){
								break;
							}
						}
						ExcelCustomValidate excelCustomValidate = excelCustomModel.getExcelModel().getExcelCustomValidate();
						ExcelRowData<?> excelRowData = new ExcelRowData<Object>() {
							@Override
							public Row getRow() {
								return excelCustomModel.getRow();
							}

							@Override
							public Cell getCell() {
								return excelCustomModel.getCell();
							}

							@Override
							public Object getRowData() {
								return finalCellData;
							}

							@Override
							public <U> U getRowData(Class<U> uClass) {
								return Reflect.mapToBean(finalCellData, uClass);
							}

							@Override
							public Object currentValue() {
								return excelCustomModel.getCurrentValue();
							}

							@Override
							public int currentCellNum() {
								return excelCustomModel.getCurrentCellNum();
							}

							@Override
							public int currentRowNum() {
								return excelCustomModel.getCurrentRowNum();
							}

							@Override
							public <U> U getOriginalCellData(Class<U> uClass) {
								return Reflect.mapToBean(finalOriginalCellData, uClass);
							}

							@Override
							public Map<String, Object> getOriginalCellData() {
								return finalOriginalCellData;
							}
						};
						boolean validate = excelCustomValidate.validate(excelRowData);
						if(!validate){
							String errMsg = "Sheet '" + sheet.getSheetName() + "': row " + (r + 1) + ", column " + (excelCustomModel.getCurrentCellNum() + 1) + " — custom validation failed: " + excelCustomValidate.errorMessage();
							errorMessage.append(errMsg).append("\n");
							if (readListener != null) readListener.onError(errMsg, r - startRow);
						}
					}
				}
				if(firstErrorBreak){
					if(errorMessage.length() > 0 ){
						return false;
					}
				}

				if (readListener != null) {
					if (errorMessage.length() == errorLengthBeforeRow) {
						readListener.onRow(cellData, r - startRow);
						listenerRowCount++;
					}
				} else {
					rowData.add(cellData);
				}
			}
			datas.add(rowData);
		}
		if (readListener != null) {
			readListener.onFinish(listenerRowCount);
		}
		if (errorMessage.length() == 0)
			return true;
		else
			return false;
	}

	private boolean checkIsEmpty(Row row) {
		try {
			short minColIx = row.getFirstCellNum();
			short maxColIx = row.getLastCellNum();
			int total = 0 ;
			int empty = 0;
			for (short colIx = minColIx; colIx <= maxColIx; colIx++) {
				total++;
				Cell cell = row.getCell(colIx);
				if (cell == null || cell.getCellType() == CellType.BLANK || getCellValue(cell) == null) {
					empty++;
				}
			}
			return total == empty;
		} catch (Exception e) {
			logger.error("get cell value error",e);
			return false;
		}
	}


	/**
	 * Retrieves a value from the map by key.
	 */
	private void setByPath(Object bean, String path, Map<String, Object> map, String valueKey) {
		String[] parts = path.split("\\.");
		Object cur = bean;
		for (int i = 0; i < parts.length - 1; i++) {
			Field pf = Reflect.findField(cur.getClass(), parts[i]);
			if (pf == null) return;
			Object next = Reflect.getField(pf, cur);
			if (next == null) {
				try {
					next = pf.getType().getDeclaredConstructor().newInstance();
				} catch (Exception e) {
					logger.debug("Cannot instantiate nested field {} on {}", parts[i], cur.getClass().getName());
					return;
				}
				setProp(cur, parts[i], pf, next);
			}
			cur = next;
		}
		String leaf = parts[parts.length - 1];
		Field lf = Reflect.findField(cur.getClass(), leaf);
		if (lf == null) return;
		Object value = getFromMap(map, valueKey, lf.getType());
		if (value != null) setProp(cur, leaf, lf, value);
	}

	private void setProp(Object owner, String name, Field field, Object value) {
		String setter = "set" + Character.toUpperCase(name.charAt(0)) + name.substring(1);
		Method m = Reflect.findMethod(owner.getClass(), setter, field.getType());
		try {
			if (m != null) {
				m.invoke(owner, value);
			} else {
				field.setAccessible(true);
				field.set(owner, value);
			}
		} catch (Exception e) {
			logger.debug("Failed to set {} on {}", name, owner.getClass().getName());
		}
	}

	private Object getFromMap(Map<?, ?> map, Object key, Class type) {
		Object value = null;
		Set<?> keys = map.keySet();
		Iterator<?> it = keys.iterator();
		while (it.hasNext()) {
			Object mapKey = it.next();
			if (key == mapKey || mapKey.equals(key.toString())) {
				value = map.get(mapKey);
				if (type != null)
					value = caseObject(value, type);

			}
		}
		return value;
	}

	/**
	 * Converts a value to the specified type.
	 */
	public Object caseObject(Object value, Class type) {
		Object result = null;
		if (value == null || value.toString().trim().length() == 0)
			return type.cast(null);
		if (type != null) {
			String className = type.getCanonicalName();
			String numVal = value.toString().replaceAll(",", "").trim();
			switch (className) {
			case "java.lang.String":
				result = value + "";
				break;
			case "java.lang.Integer":
			case "int":
				if (numVal.indexOf(".") != -1) {
					numVal = numVal.substring(0, numVal.indexOf("."));
				}
				result = Integer.valueOf(numVal);
				break;
			case "java.lang.Double":
			case "double":
				result = Double.valueOf(numVal);
				break;
			case "java.lang.Long":
			case "long":
				result = Long.valueOf(numVal);
				break;
			case "java.lang.Boolean":
			case "boolean":
				result = Boolean.valueOf(value.toString());
				break;
			case "java.lang.Float":
			case "float":
				result = Float.valueOf(numVal);
				break;
			case "java.lang.Short":
			case "short":
				result = Short.valueOf(numVal);
				break;
			case "java.util.Date":
				result = dateHandle(value.toString());
				break;
			case "java.sql.Timestamp":
				result = timeStampHandle(value.toString());
				break;
			case "java.sql.Date":
				result = sqlDateHandle(value.toString());
				break;
			case "java.math.BigDecimal":
				result = new java.math.BigDecimal(numVal);
				break;
			case "java.math.BigInteger":
				result = new java.math.BigInteger(numVal.indexOf(".") != -1
						? numVal.substring(0, numVal.indexOf("."))
						: numVal);
				break;
			case "java.lang.Character":
			case "char":
				String charStr = value.toString();
				result = charStr.isEmpty() ? null : Character.valueOf(charStr.charAt(0));
				break;
			case "java.lang.Byte":
			case "byte":
				result = Byte.valueOf(numVal);
				break;
			case "java.time.LocalDate": {
				Date d = dateHandle(value.toString());
				result = d != null ? d.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate() : null;
				break;
			}
			case "java.time.LocalDateTime": {
				Date d = dateHandle(value.toString());
				result = d != null ? d.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDateTime() : null;
				break;
			}
			}
		}
		return result;
	}
	private static final String[] DATE_PATTERNS = {
			"yyyy-MM-dd HH:mm:ss",
			"yyyy-MM-dd'T'HH:mm:ss",
			"yyyy/MM/dd HH:mm:ss",
			"yyyy-MM-dd",
			"yyyy/MM/dd",
			"yyyy.MM.dd",
			"yyyy-MM",
			"yyyy/MM",
			"yyyyMMdd"
	};

	public Date dateHandle(String source){
		if (source == null) return null;
		String trimmed = source.trim();
		for (String pattern : DATE_PATTERNS) {
			try {
				return new SimpleDateFormat(pattern).parse(trimmed);
			} catch (ParseException ignored) {}
		}
		logger.warn("Unable to parse date string '{}' with any of the known patterns", source);
		return null;
	}

	public Object timeStampHandle(String source){
		Date d = dateHandle(source);
		return d != null ? new java.sql.Timestamp(d.getTime()) : null;
	}

	public java.sql.Date sqlDateHandle(String source){
		Date d = dateHandle(source);
		return d != null ? new java.sql.Date(d.getTime()) : null;
	}

	/**
	 * Returns the parsed data from a sheet as a list of Java objects.
	 *
	 * @param sheetNum the sheet index
	 * @param clazz    the target class
	 */
	public <T> List<T> getObject(int sheetNum, Class<T> clazz) {
		LinkedList<T> objets = new LinkedList<>();
		if (sheetNum < datas.size()) {
			LinkedList<Map<String, Object>> sheetData = datas.get(sheetNum);
			if(clazz.isAssignableFrom(Map.class)) {
				return (List<T>) sheetData;
			}
			for (Map<String, Object> map : sheetData) {
				try {
					T t = clazz.getDeclaredConstructor().newInstance();
					LinkedList<ExcelModel> listModels = columnNameList.get(sheetNum);
					for (ExcelModel excelModel : listModels) {
						String field = excelModel.getFieldName();
						String sourcePath = excelModel.getSourcePath();
						// Flattened @ExcelInfoChild column: rebuild the nested object via path (keyed by sourcePath).
						if (excelModel.isFlattened() && sourcePath != null && sourcePath.contains(".")) {
							setByPath(t, sourcePath, map, sourcePath);
							continue;
						}
						Field javaField =Reflect.findField(clazz,field) ;//clazz.getDeclaredField(field);
						String setMethod = "set" + Character.toUpperCase(field.charAt(0)) + field.substring(1);
						if(javaField!=null){
							Method method = Reflect.findMethod(clazz,setMethod,javaField.getType());//clazz.getDeclaredMethod(setMethod, javaField.getType());
							if(method!=null){
								Object value = getFromMap(map, field, javaField.getType());
								if (value != null)
									method.invoke(t, value);
							}
						} else {
							logger.debug("Field '{}' not found in {}; skipping column", field, clazz.getName());
						}
					}
					objets.add(t);
				} catch (InstantiationException | IllegalAccessException | NoSuchMethodException e) {
					logger.error("Failed to instantiate target class {}", clazz.getName(), e);
				} catch (SecurityException | IllegalArgumentException e) {
					logger.error("Security or argument error mapping row to {}", clazz.getName(), e);
				} catch (InvocationTargetException e) {
					logger.error("Setter invocation failed while mapping row to {}", clazz.getName(), e);
				}
			}
		}
		return objets;
	}

	/**
	 * Returns the value of a cell.
	 */
	public Object getCellValue(Cell cell) {
		Object value = null;
		switch (cell.getCellType()) {
		case BLANK:
		case STRING:
			if(Reflect.hasText(cell.getStringCellValue())){
				value  = cell.getStringCellValue();
			}
			break;
		case BOOLEAN:
			value = cell.getBooleanCellValue();
			break;
		case ERROR:
			value = cell.getErrorCellValue();
			break;
		case FORMULA:
			value =getCellValue(formulaEvaluator.evaluateInCell(cell));
			break;
		case NUMERIC:
			value = cell.getNumericCellValue();
			value=formart(value);
			break;
		default:
			break;
		}
		return value;
	}
	
	public Object formart(Object object) {
		// NumberFormat.getInstance() silently rounds to 3 fraction digits; go through
		// BigDecimal so the full numeric precision of the cell survives the import.
		if (object instanceof Number) {
			return new java.math.BigDecimal(object.toString()).stripTrailingZeros().toPlainString();
		}
		return object.toString();
	}

	/**
	 * Extracts a picture from the Excel file at the given row/cell position and saves it locally.
	 *
	 * @param row            row index
	 * @param cell           column index
	 * @param downPath        local path to save the image
	 * @param imageVisitPrex  URL prefix for the saved image
	 * @return the access URL of the saved image, or {@code null} if no picture was found
	 */
	private String getPicture(int row, int cell, String downPath, String imageVisitPrex) {
		String value = null;
		boolean hasPicture = false;
		try {
			if (isXLS()) {
				HSSFSheet sheet = (HSSFSheet) this.sheet;
				List<HSSFPictureData> pictures = (List<HSSFPictureData>) workbook.getAllPictures();
				if (pictures.size() != 0) {
					for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
						HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
						if (shape instanceof HSSFPicture) {
							HSSFPicture pic = (HSSFPicture) shape;
							int pictureIndex = pic.getPictureIndex() - 1;
							HSSFPictureData picData = pictures.get(pictureIndex);
							int pictureRow = anchor.getRow1();
							int pictureCol = anchor.getCol1();
							if (pictureRow == row && pictureCol == cell) {
								ByteArrayInputStream input = new ByteArrayInputStream(picData.getData());
								BufferedImage image = ImageIO.read(input);
								String uuid = UUID.randomUUID().toString().replaceAll("-", "")
										+ System.currentTimeMillis();
								imageVisitPrex += uuid + "." + picData.suggestFileExtension();
								value = downPath + uuid + "." + picData.suggestFileExtension();
								try (FileOutputStream fos = new FileOutputStream(value)) {
									ImageIO.write(image, picData.suggestFileExtension(), fos);
								}
								hasPicture = true;
								break;
							}
						}
					}
				}
			}
			if (isXLSX()) {
				XSSFSheet sheet = (XSSFSheet) this.sheet;
				for (POIXMLDocumentPart dr : sheet.getRelations()) {
					if (dr instanceof XSSFDrawing) {
						XSSFDrawing drawing = (XSSFDrawing) dr;
						List<XSSFShape> shapes = drawing.getShapes();
						for (XSSFShape shape : shapes) {
							XSSFPicture pic = (XSSFPicture) shape;
							XSSFClientAnchor anchor = pic.getPreferredSize();
							CTMarker ctMarker = anchor.getFrom();
							int pictureRow = ctMarker.getRow();
							int pictureCol = ctMarker.getCol();
							if (pictureRow == row && pictureCol == cell) {
								ByteArrayInputStream input = new ByteArrayInputStream(pic.getPictureData().getData());
								BufferedImage image = ImageIO.read(input);
								String uuid = UUID.randomUUID().toString().replaceAll("-", "")
										+ System.currentTimeMillis();
								imageVisitPrex += uuid + "." + pic.getPictureData().suggestFileExtension();
								value = downPath + uuid + "." + pic.getPictureData().suggestFileExtension();
								try (FileOutputStream fos = new FileOutputStream(value)) {
									ImageIO.write(image, pic.getPictureData().suggestFileExtension(), fos);
								}
								hasPicture = true;
								break;
							}
						}
					}
				}
			}
		} catch (IOException e) {
			logger.error("Failed to extract picture at row={}, col={}", row, cell, e);
		}
		if (!hasPicture)
			imageVisitPrex = null;
		return imageVisitPrex;
	}

	/**
	 * Registers a single column index to skip during parsing.
	 *
	 * @param sheetNum    sheet index
	 * @param columnIndex column index to skip
	 */
	public void addExceptColumnNum(int sheetNum, int columnIndex) {
		if (sheetNum < 0)
			throw new IllegalArgumentException("num < 0");
		if (columnIndex < 0)
			throw new IllegalArgumentException("num < 0");
		if (exceptColumnNumMap.containsKey(sheetNum)) {
			List<Integer> child = exceptColumnNumMap.get(sheetNum);
			child.add(columnIndex);
		} else {
			List<Integer> child = new ArrayList<>();
			child.add(columnIndex);
			exceptColumnNumMap.put(sheetNum, child);
		}

	}

	/**
	 * Registers multiple column indices to skip during parsing.
	 *
	 * @param sheetNum         sheet index
	 * @param exceptColumnNums list of column indices to skip
	 */
	public void addExceptColumnNums(int sheetNum, List<Integer> exceptColumnNums) {
		if (sheetNum < 0)
			throw new IllegalArgumentException("num < 0");
		this.exceptColumnNumMap.put(sheetNum, exceptColumnNums);
	}

	public void addSheetStartRow(int startRow) {
		this.sheetStartRow.add(startRow);
	}

	public void addSheetStartRows(LinkedList<Integer> sheetStartRow) {
		this.sheetStartRow.addAll(sheetStartRow);
	}

	/**
	 * Sets the row index from which data reading starts.
	 */
	public void setReadListener(ExcelReadListener readListener) {
		this.readListener = readListener;
	}

	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}

	/**
	 * Adds a column model list for one sheet.
	 *
	 * @param excelModels column models for the sheet
	 */
	public void addColumnName(LinkedList<ExcelModel> excelModels) {
		this.columnNameList.add(excelModels);
	}

	/**
	 * Adds column model lists for all sheets.
	 */
	public void addColumnNames(LinkedList<LinkedList<ExcelModel>> excelModels) {
		this.columnNameList.addAll(excelModels);
	}

	/**
	 * Creates an ExcelModel for a field with full configuration.
	 *
	 * @param fieldName            field name
	 * @param needtranslate        whether the field requires value translation
	 * @param translateMappingInfo translation mapping
	 * @param isPicture            whether the field is a picture
	 * @param imageDownPath        local path to download the image to
	 * @param imageVisitPrex       URL prefix for image access
	 */
	public static ExcelModel generateExcelModel(String fieldName, boolean needtranslate,
			Map<Object, Object> translateMappingInfo, boolean isPicture, String imageDownPath, String imageVisitPrex) {
		return new ExcelModel(fieldName, needtranslate, translateMappingInfo, isPicture, imageDownPath, imageVisitPrex);
	}

	/**
	 * Creates a basic ExcelModel for a plain field.
	 *
	 * @param fieldName field name
	 */
	public static ExcelModel generateExcelModel(String fieldName) {
		return new ExcelModel(fieldName);
	}

	public static ExcelModel generateExcelModel(String fieldName,boolean nullAble) {
		return new ExcelModel(fieldName,nullAble);
	}

	/**
	 * Creates an ExcelModel with a translation mapping.
	 *
	 * @param fieldName            field name
	 * @param translateMappingInfo translation mapping
	 */
	public static ExcelModel generateExcelModel(String fieldName, Map<Object, Object> translateMappingInfo) {
		return new ExcelModel(fieldName, translateMappingInfo);
	}

	/**
	 * Creates an ExcelModel for a picture field.
	 *
	 * @param fieldName      field name
	 * @param imageVisitPrex URL prefix for image access
	 * @param imageDownPath  local path to download the image to
	 */
	public static ExcelModel generateExcelModel(String fieldName, String imageVisitPrex, String imageDownPath) {
		return new ExcelModel(fieldName, imageVisitPrex, imageDownPath);
	}

	/**
	 * Creates an ExcelModel for a date field.
	 *
	 * @param fieldName field name
	 * @param pattern   date format pattern; defaults to {@code yyyy-MM-dd HH:mm:ss}
	 */
	public static ExcelModel generateExcelModel(String fieldName, String pattern) {
		return new ExcelModel(fieldName, pattern);
	}

	/**
	 * Returns the number of sheets in the current workbook.
	 */
	public int getSheetNums(){
		return workbook.getNumberOfSheets();
	}

	/**
	 * Returns the current workbook instance.
	 */
	public  Workbook getWorkbook(){
		return  workbook;
	}
}
