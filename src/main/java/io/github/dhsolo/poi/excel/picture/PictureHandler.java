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
package io.github.dhsolo.poi.excel.picture;

import io.github.dhsolo.poi.excel.ExcelModel;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Strategy interface for picture processing in Excel generation.
 *
 * <p>Handles image downloading, resizing, embedding, and temporary file management.
 * Implemented by {@link DefaultPictureHandler}.
 *
 * @author dhsolo
 * @since 1.0
 */
public interface PictureHandler {

    /**
     * Creates the temporary directory used to store downloaded images before they are
     * embedded into the workbook. Must be called before {@link #downLoadPicture()}.
     */
    void createTempleFileDir();

    /**
     * Pre-scans the data list to determine the maximum number of images in any single
     * picture column cell, then updates the header array if column expansion is required.
     *
     * @param columnMappingInfo   column-index-to-{@link ExcelModel} map for the sheet
     * @param dataList            the full list of row data objects to be exported
     * @param header              current header label array; may be expanded in place
     * @param needOrderNum        {@code true} when a leading order-number column is present
     * @param orderColumnSpan     number of columns the order-number column spans (ignored when
     *                            {@code needOrderNum} is {@code false})
     */
    void checkPictureMaxSize(Map<Integer, ExcelModel> columnMappingInfo, List<?> dataList,
                             String[] header, boolean needOrderNum, int orderColumnSpan);

    /**
     * Expands the header array to accommodate multi-image columns by inserting blank
     * header labels for any extra columns reserved for overflow images.
     *
     * @param header the original header label array
     * @return a new array with additional blank entries inserted after each expanded picture column
     */
    String[] expandHeaderForPictures(String[] header);

    /**
     * Concurrently downloads all images whose URLs were registered via
     * {@link #setPicture} into the temporary directory. Must be called after the
     * sheet rendering pass has completed.
     */
    void downLoadPicture();

    /**
     * Embeds a single image into the workbook at the specified cell region and registers
     * a download task for the image URL.
     *
     * @param startColumn 0-based starting column index of the image anchor
     * @param startRow    0-based starting row index of the image anchor
     * @param endColumn   0-based ending column index of the image anchor
     * @param endRow      0-based ending row index of the image anchor
     * @param imageUrl    URL of the image to download and embed; {@code null} skips embedding
     * @param cell        the cell at the anchor position (used for workbook reference)
     * @return the number of extra columns consumed when the image spans more than one column
     */
    int setPicture(int startColumn, int startRow, int endColumn, int endRow, String imageUrl, Cell cell);

    /**
     * Decompresses the downloaded image archive (if zipped), embeds all images into the
     * workbook, and re-compresses the workbook as a ZIP file for streaming. Called as the
     * final step of the picture pipeline.
     */
    void decompressionPictureDirAndCompression();

    /**
     * Deletes all temporary files and directories created during picture processing.
     * Should be called after the workbook has been exported.
     */
    void cleanup();

    // --- State mutators ---

    /** Sets the POI picture type constant (e.g., {@code Workbook.PICTURE_TYPE_JPEG}). */
    void setPictureType(int pictureType);

    /** Sets the HTTP read timeout in milliseconds for image downloads. */
    void setImageReadTimeOut(int imageReadTimeOut);

    /**
     * Indicates whether the sheet uses a child-complex (nested header) layout, which
     * affects row/column anchor calculations during image embedding.
     */
    void setChildComplex(boolean childComplex);

    /** Sets the sequential image counter used to generate unique anchor names. */
    void setImageIndex(int imageIndex);

    /** Marks whether at least one picture column has been registered. */
    void setHasPicture(boolean hasPicture);

    /** Marks whether the temporary image directory has been created on disk. */
    void setPictureDirCreate(boolean pictureDirCreate);

    /** Sets the path of the directory where downloaded images are stored. */
    void setCurrentPictureDownLoadDir(String dir);

    /**
     * Controls whether downloaded images should be resized before embedding.
     * When {@code true}, {@link #setResizeWith} and {@link #setResizeHeight} must also be set.
     */
    void setNeedImageResize(boolean needImageResize);

    /** Sets the target width in pixels for image resizing. */
    void setResizeWith(int resizeWith);

    /** Sets the target height in pixels for image resizing. */
    void setResizeHeight(int resizeHeight);

    // --- State accessors ---

    /** Returns the current sequential image counter. */
    int getImageIndex();

    /** Returns {@code true} if at least one picture column URL has been registered. */
    boolean hasPicture();

    /**
     * Returns the mapping from column index to the maximum number of images found in
     * any single cell of that column. Used to determine how many extra columns must be
     * inserted for overflow images.
     *
     * @return column-index → max-image-count map
     */
    Map<Integer, Integer> getColumnMaxMapping();

    /**
     * Returns the column-expansion mapping computed by the most recent
     * {@link #checkPictureMaxSize} call only. On a complex multi-section sheet the handler is
     * shared and {@link #getColumnMaxMapping()} accumulates entries from every section, whose
     * data-column key spaces collide; per-section layout decisions must use this view instead.
     *
     * @return column-index → max-extra-image-count map for the last analysed section
     */
    default Map<Integer, Integer> getSectionColumnMaxMapping() {
        return getColumnMaxMapping();
    }

    /**
     * Variant of {@link #expandHeaderForPictures(String[])} driven by an explicit expansion
     * mapping (typically {@link #getSectionColumnMaxMapping()}), so a section of a complex
     * sheet is not expanded by another section's entries.
     *
     * @param header    the header labels to expand
     * @param columnMax column-index → max-extra-image-count map to apply
     * @return the expanded header (the same array when no expansion applies)
     */
    default String[] expandHeaderForPictures(String[] header, Map<Integer, Integer> columnMax) {
        return expandHeaderForPictures(header);
    }

    /** Returns the filesystem path of the temporary image download directory. */
    String getCurrentPictureDownLoadDir();

    /** Returns {@code true} if the temporary image directory has been created on disk. */
    boolean isPictureDirCreate();

    /**
     * Returns the temporary workbook file created during ZIP compression, or {@code null}
     * when no picture processing has occurred.
     *
     * @return temp workbook file, may be {@code null}
     */
    File getTempWorkFile();

    /**
     * Returns {@code true} when the workbook was re-packaged as a ZIP archive after image
     * embedding, meaning the final output must be read from {@link #getTempWorkFile()}.
     */
    boolean isZip();

    /**
     * Returns all image download tasks registered during the sheet render pass.
     * These tasks are executed by {@link #downLoadPicture()}.
     *
     * @return list of pending download tasks
     */
    List<ImageDownLoadTask> getImageDownLoadTasks();
}
