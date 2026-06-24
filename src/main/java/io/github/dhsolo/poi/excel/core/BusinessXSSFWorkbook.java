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

import java.lang.reflect.Field;
import java.util.List;

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
 * @author dhsolo
 * @since 1.0
 */
public class BusinessXSSFWorkbook extends XSSFWorkbook {

    /**
     * Registers a JPEG image relationship at the given index and adds the corresponding
     * {@link XSSFPictureData} to the workbook's picture list (skipping parts the lazy
     * {@link #getAllPictures()} scan already registered).
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
        // Anchors created during data population resolve pictures by index into the workbook's
        // picture list, so the part created above must be registered there. The list returned
        // by getAllPictures() is unmodifiable since POI 5.4, so append via the internal field;
        // getAllPictures() is called first to trigger the lazy scan, which — when the list was
        // never initialised — already wraps this part in a NEW XSSFPictureData instance.
        // XSSFPictureData has no equals override, so dedup must compare package part names,
        // not wrapper identity: an identity contains() would re-add the part and shift every
        // later by-index anchor to the wrong image.
        getAllPictures();
        List<XSSFPictureData> internal = internalPictures();
        var partName = img.getPackagePart().getPartName();
        boolean registered = false;
        for (XSSFPictureData existing : internal) {
            if (partName.equals(existing.getPackagePart().getPartName())) {
                registered = true;
                break;
            }
        }
        if (!registered) {
            internal.add(img);
        }
    }

    /**
     * The {@code XSSFWorkbook.pictures} field, resolved once at class load so a POI upgrade
     * that renames it fails fast (and loudly) instead of mid-export.
     */
    private static final Field PICTURES_FIELD;
    static {
        try {
            PICTURES_FIELD = XSSFWorkbook.class.getDeclaredField("pictures");
            PICTURES_FIELD.setAccessible(true);
        } catch (ReflectiveOperationException e) {
            throw new ExceptionInInitializerError(new IllegalStateException(
                    "Cannot access XSSFWorkbook.pictures; POI internals changed — picture pre-registration needs porting", e));
        }
    }

    /** The workbook's internal (mutable) picture list. */
    @SuppressWarnings("unchecked")
    private List<XSSFPictureData> internalPictures() {
        try {
            return (List<XSSFPictureData>) PICTURES_FIELD.get(this);
        } catch (IllegalAccessException e) {
            throw new IllegalStateException(
                    "Cannot access XSSFWorkbook.pictures; POI internals changed — picture pre-registration needs porting", e);
        }
    }
}
