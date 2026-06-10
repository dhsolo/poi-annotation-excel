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
import io.github.dhsolo.poi.excel.core.ValueExtractor;
import io.github.dhsolo.poi.excel.core.BusinessSXSSFWorkbook;
import io.github.dhsolo.poi.excel.core.BusinessXSSFWorkbook;
import io.github.dhsolo.poi.excel.exception.ExcelException;
import io.github.dhsolo.zip.ZipUtil;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.regex.Pattern;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Default implementation of {@link PictureHandler}.
 *
 * <p>Handles image downloading, resizing, embedding, and temporary file management.
 *
 * @author dhsolo
 * @since 1.0
 */
public class DefaultPictureHandler implements PictureHandler {

    private static final Logger logger = LoggerFactory.getLogger(DefaultPictureHandler.class);

    static final int MOVE_AND_RESIZE = 0;
    static final int MOVE_DONT_RESIZE = 2;
    static final int DONT_MOVE_AND_RESIZE = 3;
    static final String IMAGE_SPLIT = ",";
    private static final int DEFAULT_IMAGE_READ_TIME_OUT = 2000;
    private static final String TEMP_DIR_PROPERTY = "java.io.tmpdir";
    private static final int BATCH_SIZE = 50;

    private final Workbook book;
    private final String currentExcelType;
    private final Sheet sheet;
    private final Drawing drawing;
    private final String imagesSeparator;
    /** Pre-compiled split pattern for {@link #imagesSeparator}, reused across all cells. */
    private final Pattern imagesSeparatorPattern;

    private int pictureType = MOVE_AND_RESIZE;
    private int imageReadTimeOut = DEFAULT_IMAGE_READ_TIME_OUT;
    private boolean needImageResize = false;
    private int resizeWith;
    private int resizeHeight;

    private String currentPictureDownLoadDir;
    private boolean pictureDirCreate = false;
    private boolean hasPicture = false;
    private boolean isChildComplex;

    private final AtomicInteger imageNum = new AtomicInteger(0);
    private int imageIndex = 1;
    private final Map<Integer, Integer> columnMaxMapping = new HashMap<>();
    private final List<ImageDownLoadTask> imageDownLoadTasks = new ArrayList<>();

    /** Futures for submitted download tasks; joined at the pre-injection barrier. */
    private final List<Future<?>> downloadFutures = new ArrayList<>();
    /** Number of tasks already submitted, so parent and child passes do not resubmit. */
    private int submittedTaskCount = 0;
    /** Guards {@link #awaitDownloads()} against being run more than once. */
    private boolean downloadsAwaited = false;

    private File tempWorkFile;
    private boolean isZip = false;

    private final ThreadPoolExecutor executor;
    private final ValueExtractor valueExtractor;

    public DefaultPictureHandler(Workbook book, String currentExcelType, Sheet sheet, Drawing drawing,
                        String imagesSeparator, ThreadPoolExecutor executor, ValueExtractor valueExtractor) {
        this.book = book;
        this.currentExcelType = currentExcelType;
        this.sheet = sheet;
        this.drawing = drawing;
        this.imagesSeparator = imagesSeparator;
        this.imagesSeparatorPattern = Pattern.compile(imagesSeparator != null ? imagesSeparator : ",");
        this.executor = executor;
        this.valueExtractor = valueExtractor;
    }

    @Override
    public void setPictureType(int pictureType) {
        this.pictureType = pictureType;
    }

    @Override
    public void setImageReadTimeOut(int imageReadTimeOut) {
        this.imageReadTimeOut = imageReadTimeOut;
    }

    @Override
    public void setNeedImageResize(boolean needImageResize) {
        this.needImageResize = needImageResize;
    }

    @Override
    public void setResizeWith(int resizeWith) {
        this.resizeWith = resizeWith;
    }

    @Override
    public void setResizeHeight(int resizeHeight) {
        this.resizeHeight = resizeHeight;
    }

    @Override
    public void setChildComplex(boolean childComplex) {
        isChildComplex = childComplex;
    }

    @Override
    public void setImageIndex(int imageIndex) {
        this.imageIndex = imageIndex;
    }

    @Override
    public int getImageIndex() {
        return imageIndex;
    }

    @Override
    public boolean hasPicture() {
        return hasPicture;
    }

    @Override
    public Map<Integer, Integer> getColumnMaxMapping() {
        return columnMaxMapping;
    }

    @Override
    public String getCurrentPictureDownLoadDir() {
        return currentPictureDownLoadDir;
    }

    @Override
    public boolean isPictureDirCreate() {
        return pictureDirCreate;
    }

    @Override
    public File getTempWorkFile() {
        return tempWorkFile;
    }

    @Override
    public boolean isZip() {
        return isZip;
    }

    @Override
    public List<ImageDownLoadTask> getImageDownLoadTasks() {
        return imageDownLoadTasks;
    }

    @Override
    public void setPictureDirCreate(boolean pictureDirCreate) {
        this.pictureDirCreate = pictureDirCreate;
    }

    @Override
    public void setCurrentPictureDownLoadDir(String dir) {
        this.currentPictureDownLoadDir = dir;
    }

    @Override
    public void setHasPicture(boolean hasPicture) {
        this.hasPicture = hasPicture;
    }

    /**
     * Creates temporary directory for picture downloads.
     */
    @Override
    public void createTempleFileDir() {
        currentPictureDownLoadDir = System.getProperty(TEMP_DIR_PROPERTY);
        currentPictureDownLoadDir += File.separator + java.util.UUID.randomUUID();
        logger.debug("picture downLoad dir :" + currentPictureDownLoadDir);
        File file = new File(currentPictureDownLoadDir);
        if (!file.exists()) {
            try {
                boolean mkdirs = file.mkdirs();
                pictureDirCreate = mkdirs;
                if (!mkdirs) {
                    logger.debug("picture downLoad dir create false");
                }
            } catch (Exception e) {
                logger.error("picture downLoad dir create false", e);
            }
        } else {
            pictureDirCreate = true;
        }
    }

    /**
     * Checks max picture column count and prepares download tasks.
     */
    @Override
    public void checkPictureMaxSize(Map<Integer, ExcelModel> columnMappingInfo, List<?> dataList,
                             String[] header, boolean needOrderNum, int orderColumnSpan) {
        int size = columnMappingInfo.size();
        Map<Integer, ExcelModel> pictureExcelModel = new LinkedHashMap<>();
        for (int i = 0; i < size; i++) {
            ExcelModel excelModel = columnMappingInfo.get(i);
            if (excelModel.isPicture()) {
                pictureExcelModel.put(i, excelModel);
            }
        }
        if (pictureExcelModel.isEmpty()) return;

        int indexNum = imageIndex;
        Map<Integer, String> pictureIndexMapping = new LinkedHashMap<>();
        // Full index→url map retained so each pre-created relationship can be typed per format.
        Map<Integer, String> allIndexUrl = new LinkedHashMap<>();
        int totalNum = 0;
        int imageTaskNum = 1;

        for (Object obj : dataList) {
            for (Map.Entry<Integer, ExcelModel> entry : pictureExcelModel.entrySet()) {
                Integer key = entry.getKey();
                ExcelModel value = entry.getValue();
                String fieldName = value.getFieldName();
                Object picture = valueExtractor.getValue(fieldName, obj);
                if (picture != null && picture.toString().trim().length() > 0) {
                    String s = picture.toString().trim();
                    String[] split = imagesSeparatorPattern.split(s);
                    for (String url : split) {
                        int curr = indexNum++;
                        pictureIndexMapping.put(curr, url);
                        allIndexUrl.put(curr, url);
                        totalNum++;
                    }
                    int length = split.length - 1;
                    if (length > 0) {
                        Integer max = columnMaxMapping.computeIfAbsent(key, k -> length);
                        if (length > max) {
                            columnMaxMapping.put(key, length);
                        }
                    }
                }
            }

            if (pictureIndexMapping.size() >= BATCH_SIZE) {
                imageDownLoadTasks.add(new ImageDownLoadTask(pictureIndexMapping, currentPictureDownLoadDir,
                        imageReadTimeOut, needImageResize, resizeWith, resizeHeight, imageTaskNum++));
                pictureIndexMapping = new LinkedHashMap<>();
            }
        }

        imageIndex = indexNum;
        // Pre-create one image relationship per index, typed by the URL's format, so that
        // anchors created during data population resolve correctly and the bytes downloaded to
        // disk can be injected straight into the matching xl/media/image#.<ext> part.
        for (Map.Entry<Integer, String> e : allIndexUrl.entrySet()) {
            PictureFormat fmt = PictureFormat.fromUrl(e.getValue());
            if (book instanceof BusinessXSSFWorkbook xx) {
                xx.addPicture(e.getKey(), fmt.relation());
            } else if (book instanceof BusinessSXSSFWorkbook sx) {
                sx.addPicture(e.getKey(), fmt.relation());
            }
        }

        if (!pictureIndexMapping.isEmpty()) {
            imageDownLoadTasks.add(new ImageDownLoadTask(pictureIndexMapping, currentPictureDownLoadDir,
                    imageReadTimeOut, needImageResize, resizeWith, resizeHeight, imageTaskNum++));
        }

        // Activate the disk-staging path: with images present, downloads go to disk and are
        // injected into the ZIP at the end, keeping image bytes off the heap.
        if (totalNum > 0) {
            hasPicture = true;
        }

        logger.debug("picture task num : {}, picture num : {}", imageDownLoadTasks.size(), totalNum);
    }

    /**
     * Expands header array for multi-picture columns.
     */
    @Override
    public String[] expandHeaderForPictures(String[] header) {
        if (columnMaxMapping.isEmpty()) {
            return header;
        }
        List<String> newHeader = new ArrayList<>();
        for (int i = 0; i < header.length; i++) {
            Integer maxNum = columnMaxMapping.get(i);
            String headName = header[i];
            newHeader.add(headName);
            if (maxNum != null) {
                while (maxNum-- > 0) {
                    newHeader.add(headName);
                }
            }
        }
        return newHeader.toArray(new String[0]);
    }

    /**
     * Submits image-download tasks to the thread pool and returns immediately, so that the
     * caller can populate workbook data in parallel while images download to disk.
     *
     * <p>Only tasks not yet submitted are dispatched, allowing parent and child sheets (which
     * share this handler) to each contribute tasks without resubmitting. Completion is awaited
     * later at {@link #awaitDownloads()}, just before the images are injected into the ZIP.
     */
    @Override
    public void downLoadPicture() {
        if (imageDownLoadTasks == null || imageDownLoadTasks.isEmpty()) {
            logger.debug("No image download tasks to process");
            return;
        }
        if (!(pictureDirCreate && hasPicture)) {
            return;
        }

        for (int i = submittedTaskCount; i < imageDownLoadTasks.size(); i++) {
            ImageDownLoadTask task = imageDownLoadTasks.get(i);
            if (task != null) {
                downloadFutures.add(executor.submit(task));
            }
        }
        submittedTaskCount = imageDownLoadTasks.size();
        logger.debug("submitted download tasks async, pending futures={}", downloadFutures.size());
    }

    /**
     * Synchronisation barrier: blocks until every submitted download task has finished.
     * Invoked just before downloaded images are injected into the workbook archive, so the
     * download phase overlaps data population but always completes before injection.
     */
    private void awaitDownloads() {
        if (downloadsAwaited) {
            return;
        }
        downloadsAwaited = true;
        for (Future<?> future : downloadFutures) {
            if (future == null) {
                continue;
            }
            try {
                future.get(60, TimeUnit.SECONDS);
            } catch (TimeoutException e) {
                logger.warn("Image download task timeout", e);
                cancelFutureSafely(future);
            } catch (InterruptedException e) {
                logger.warn("Image download task interrupted", e);
                Thread.currentThread().interrupt();
                cancelFutureSafely(future);
            } catch (ExecutionException e) {
                logger.error("Image download task execution error", e);
                cancelFutureSafely(future);
            } catch (Exception e) {
                logger.error("Unexpected error in image download", e);
                cancelFutureSafely(future);
            }
        }
        logger.debug("Image download complete");
    }

    private static ClientAnchor.AnchorType toAnchorType(int id) {
        return switch (id) {
            case 2 -> ClientAnchor.AnchorType.MOVE_DONT_RESIZE;
            case 3 -> ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE;
            default -> ClientAnchor.AnchorType.MOVE_AND_RESIZE;
        };
    }

    private void cancelFutureSafely(Future<?> future) {
        if (future != null && !future.isCancelled()) {
            try {
                future.cancel(true);
            } catch (Exception e) {
                logger.warn("Failed to cancel future task", e);
            }
        }
    }

    /**
     * Embeds picture into the specified cell.
     *
     * @return the number of additional cells consumed by pictures
     */
    @Override
    public int setPicture(int startColumn, int startRow, int endColumn, int endRow, String imageUrl, Cell cell) {
        if (imageUrl == null || imageUrl.trim().length() == 0) {
            return 0;
        }
        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
        BufferedImage bufferImg;
        String[] imageArray = imagesSeparatorPattern.split(imageUrl);
        int count = 0;
        for (String s : imageArray) {
            ClientAnchor anchor;
            if (currentExcelType.equals("xlsx")) {
                anchor = new org.apache.poi.xssf.usermodel.XSSFClientAnchor(0, 0, 0, 0, (short) startColumn, startRow, (short) endColumn + 1,
                        endRow + 1);
            } else {
                anchor = new org.apache.poi.hssf.usermodel.HSSFClientAnchor(0, 0, 0, 0, (short) startColumn, startRow, (short) (endColumn + 1),
                        endRow + 1);
            }
            ClientAnchor.AnchorType resolved = toAnchorType(pictureType);
            anchor.setAnchorType(resolved != null ? resolved : ClientAnchor.AnchorType.MOVE_AND_RESIZE);
            int i;
            if (pictureDirCreate && hasPicture) {
                i = imageNum.getAndIncrement();
            } else {
                try {
                    byteArrayOut.reset();
                    // Same guards as the async path: protocol whitelist, ImageDownloadPolicy
                    // (SSRF), timeouts and the size cap — this fallback used to bypass them all.
                    try (java.io.InputStream is = ImageDownLoadTask.openGuardedStream(s, imageReadTimeOut)) {
                        bufferImg = ImageIO.read(ImageIO.createImageInputStream(is));
                    }
                    if (bufferImg == null) {
                        continue;
                    }
                    ImageIO.write(bufferImg, "jpg", byteArrayOut);
                    bufferImg.flush();
                } catch (Exception e) {
                    logger.error("Image download failed");
                    cell.setCellValue("Image download failed");
                    continue;
                }
                i = book.addPicture(byteArrayOut.toByteArray(), Workbook.PICTURE_TYPE_JPEG);
            }
            drawing.createPicture(anchor, i);
            count++;
            startColumn += 1;
            endColumn += 1;
        }
        return count == 0 ? count : count - 1;
    }

    /**
     * Writes the workbook to a temporary xlsx file and injects downloaded
     * pictures directly into the ZIP structure via NIO ZipFileSystem.
     *
     * <p>This avoids the costly unzip→copy→rezip cycle by operating on the
     * ZIP archive in place.
     */
    @Override
    public void decompressionPictureDirAndCompression() {
        if (!pictureDirCreate || !hasPicture) return;

        // Barrier: ensure all parallel downloads have finished before injecting their files.
        awaitDownloads();

        String tempDir = System.getProperty(TEMP_DIR_PROPERTY) + File.separator + java.util.UUID.randomUUID() + "." + currentExcelType;
        logger.debug("temp Excel File ：{}", tempDir);
        tempWorkFile = new File(tempDir);

        try {
            // Write and fully close the workbook stream BEFORE injecting: the ZipFileSystem
            // re-opens and rewrites the same file on close, which fails on Windows if an
            // output stream still holds the file open.
            try (FileOutputStream fo = new FileOutputStream(tempWorkFile)) {
                book.write(fo);
            }

            File pictureDir = new File(currentPictureDownLoadDir);
            File[] files = pictureDir.listFiles();
            if (files != null && files.length > 0) {
                ZipUtil.injectFiles(tempWorkFile.toPath(), files, "xl/media");
            }

            isZip = true;
            deleteDirectory(pictureDir.toPath());
            logger.debug("decompressionPictureDirAndCompression finished, {} pictures injected", files != null ? files.length : 0);
        } catch (IOException e) {
            throw new ExcelException("Failed to create temp Excel file", e);
        }
    }

    /**
     * Cleans up temporary files.
     */
    @Override
    public void cleanup() {
        try {
            if (tempWorkFile != null && tempWorkFile.exists()) {
                if (tempWorkFile.delete()) {
                    logger.debug("Temporary work file deleted: {}", tempWorkFile.getAbsolutePath());
                }
            }
            if (currentPictureDownLoadDir != null) {
                File pictureDir = new File(currentPictureDownLoadDir);
                if (pictureDir.exists()) {
                    deleteDirectory(pictureDir.toPath());
                    logger.debug("Picture download directory cleaned: {}", currentPictureDownLoadDir);
                }
            }
        } catch (Exception e) {
            logger.warn("Failed to cleanup temporary files", e);
        }
    }

    private static void deleteDirectory(Path dir) throws IOException {
        if (!Files.exists(dir)) return;
        try (var stream = Files.walk(dir)) {
            stream.sorted(java.util.Comparator.reverseOrder())
                  .map(Path::toFile)
                  .forEach(File::delete);
        }
    }
}
