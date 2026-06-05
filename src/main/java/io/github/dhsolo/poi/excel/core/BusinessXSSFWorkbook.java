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

import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Business-level extension of {@link XSSFWorkbook} that adds support for registering
 * embedded JPEG pictures by relationship index.
 *
 * <p>The standard {@link XSSFWorkbook} API does not expose a straightforward way to
 * attach pre-existing image relationships by index. This subclass calls the protected
 * {@code createRelationship} method to create or reuse a JPEG image part and then
 * registers it in the workbook's picture list so that downstream rendering code can
 * reference the image by its index.</p>
 *
 * <p>Used directly by {@link XlsxWorkbookStrategy} and as the underlying XSSF workbook
 * wrapped by {@link BusinessSXSSFWorkbook} in big-data mode.</p>
 *
 * @author dh
 * @since 1.0
 */
public class BusinessXSSFWorkbook extends XSSFWorkbook {

    /**
     * Registers a JPEG image relationship at the given index and adds the corresponding
     * {@link XSSFPictureData} to the workbook's picture list.
     *
     * <p>When {@code index} is {@code 1} the picture list is only initialised (via
     * {@link #getAllPictures()}); for any other index the new picture part is appended to
     * the list.</p>
     *
     * @param index the one-based relationship index used to create or locate the JPEG
     *              image part within the workbook package
     */
    public void addPicture(int index){
        addPicture(index, XSSFRelation.IMAGE_JPEG);
    }

    /**
     * Registers an image relationship of the given type at the given index and adds the
     * corresponding {@link XSSFPictureData} to the workbook's picture list.
     *
     * <p>The {@code relation} determines the part's content type and the file extension of
     * its {@code xl/media/image#.*} name, so that pre-created relationships line up with the
     * image files later injected into the package.</p>
     *
     * @param index    the one-based relationship index used to create or locate the image part
     * @param relation the POI relationship describing the image format (e.g.
     *                 {@link XSSFRelation#IMAGE_PNG})
     */
    public void addPicture(int index, XSSFRelation relation){
        var rp = createRelationship(relation, XSSFFactory.getInstance(), index, true);
        XSSFPictureData img = rp.getDocumentPart();
        if(index == 1){
            getAllPictures();
        }else {
            getAllPictures().add(img);
        }
    }
}
