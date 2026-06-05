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
package io.github.dhsolo.poi.excel.core;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * Business-level extension of {@link SXSSFWorkbook} that produces {@link BusinessSXSSFSheet}
 * instances and delegates picture management to the underlying {@link BusinessXSSFWorkbook}.
 *
 * <p>{@link SXSSFWorkbook} internally creates plain {@link SXSSFSheet} objects and registers
 * them via the package-private {@code registerSheetMapping} method. This class overrides
 * {@link #createSheet()} and {@link #createSheet(String)} so that every sheet created through
 * the workbook is a {@link BusinessSXSSFSheet} — which correctly tracks the last written row
 * number. The internal sheet-to-XSSF mapping is maintained by reflectively invoking
 * {@code registerSheetMapping}, because that method is not part of the public API.</p>
 *
 * <p>Used by {@link BigDataWorkbookStrategy}.</p>
 *
 * @author dh
 * @since 1.0
 * @see BusinessSXSSFSheet
 * @see BusinessXSSFWorkbook
 */
public class BusinessSXSSFWorkbook extends SXSSFWorkbook {

    /**
     * Creates a streaming workbook backed by the given {@link BusinessXSSFWorkbook}.
     *
     * @param xssfWorkbook the underlying XSSF workbook that stores the workbook metadata
     *                     and is used for picture management
     */
    public BusinessSXSSFWorkbook(BusinessXSSFWorkbook xssfWorkbook) {
        super(xssfWorkbook);
    }

    /**
     * Delegates picture registration to the underlying {@link BusinessXSSFWorkbook}.
     *
     * @param index the one-based relationship index of the JPEG image to register
     */
    public void addPicture(int index){
        ((BusinessXSSFWorkbook)getXSSFWorkbook()).addPicture(index);
    }

    /**
     * Delegates picture registration to the underlying {@link BusinessXSSFWorkbook} using the
     * given image relationship type.
     *
     * @param index    the one-based relationship index of the image to register
     * @param relation the POI relationship describing the image format
     */
    public void addPicture(int index, org.apache.poi.xssf.usermodel.XSSFRelation relation){
        ((BusinessXSSFWorkbook)getXSSFWorkbook()).addPicture(index, relation);
    }

    /**
     * Creates a new sheet with a default name and registers it as a {@link BusinessSXSSFSheet}.
     *
     * @return a new {@link BusinessSXSSFSheet} added to this workbook
     * @throws RuntimeException if sheet creation or internal mapping registration fails
     */
    @Override
    public SXSSFSheet createSheet() {
        return createAndRegisterSXSSFSheet(getXSSFWorkbook().createSheet());
    }

    /**
     * Creates a new named sheet and registers it as a {@link BusinessSXSSFSheet}.
     *
     * @param sheetname the name for the new sheet; must comply with Excel sheet-name rules
     * @return a new {@link BusinessSXSSFSheet} added to this workbook
     * @throws RuntimeException if sheet creation or internal mapping registration fails
     */
    @Override
    public SXSSFSheet createSheet(String sheetname) {
        return createAndRegisterSXSSFSheet(getXSSFWorkbook().createSheet(sheetname));
    }

    /**
     * Constructs a {@link BusinessSXSSFSheet} for the given {@link XSSFSheet} and registers
     * the streaming-to-XSSF mapping via reflection, because
     * {@code SXSSFWorkbook.registerSheetMapping} is package-private.
     *
     * @param xSheet the underlying XSSF sheet to wrap
     * @return the newly created and registered {@link BusinessSXSSFSheet}
     * @throws RuntimeException wrapping {@link IOException} if the sheet buffer cannot be
     *                          initialised, or wrapping reflection exceptions if the mapping
     *                          method cannot be accessed or invoked
     */
    private SXSSFSheet createAndRegisterSXSSFSheet(XSSFSheet xSheet){
        BusinessSXSSFSheet sheet;
        try{
            sheet = new BusinessSXSSFSheet(this, xSheet);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        try {
            Method registerSheetMapping = SXSSFWorkbook.class.getDeclaredMethod("registerSheetMapping", SXSSFSheet.class, XSSFSheet.class);
            registerSheetMapping.setAccessible(true);
            registerSheetMapping.invoke(this, sheet, xSheet);
        } catch (NoSuchMethodException | InvocationTargetException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
        return sheet;
    }
}
