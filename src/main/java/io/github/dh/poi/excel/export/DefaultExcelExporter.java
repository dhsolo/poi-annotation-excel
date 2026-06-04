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
package io.github.dh.poi.excel.export;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

/**
 * Default implementation of {@link ExcelExporter}.
 *
 * <p>Exports Excel workbooks to output streams, local files, or input streams.
 * Supports zip-compressed temporary files when pictures are included.
 *
 * @author dh
 * @since 1.0
 */
public class DefaultExcelExporter implements ExcelExporter {

    private static final Logger logger = LoggerFactory.getLogger(DefaultExcelExporter.class);

    private final Workbook book;
    private final String currentExcelType;
    private File tempWorkFile;
    private boolean isZip;
    private boolean isReCreate;

    /**
     * Creates an exporter bound to the given workbook.
     *
     * @param book              the fully-built POI workbook to export
     * @param currentExcelType  file extension without the leading dot, e.g. {@code "xlsx"}
     *                          or {@code "xls"}; appended by {@link #exportLocal} when missing
     */
    public DefaultExcelExporter(Workbook book, String currentExcelType) {
        this.book = book;
        this.currentExcelType = currentExcelType;
    }

    @Override
    public void setTempWorkFile(File tempWorkFile) {
        this.tempWorkFile = tempWorkFile;
    }

    @Override
    public void setZip(boolean zip) {
        isZip = zip;
    }

    @Override
    public Workbook getWorkBook() {
        if (isZip && !isReCreate) {
            try (FileInputStream fis = new FileInputStream(tempWorkFile)) {
                Workbook recreated = WorkbookFactory.create(fis);
                isReCreate = true;
                return recreated;
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
        return book;
    }

    @Override
    public void export(OutputStream outputStream, String exportFileName) throws IOException {
        try {
            if (isZip && !isReCreate) {
                try (FileInputStream fis = new FileInputStream(tempWorkFile)) {
                    fis.transferTo(outputStream);
                }
            } else {
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                book.write(baos);
                outputStream.write(baos.toByteArray());
            }
            outputStream.flush();
        } finally {
            deleteTempFile();
        }
    }

    @Override
    public <R> R upload(ExcelUploader<R> uploader, String exportFileName) throws IOException {
        byte[] data;
        try {
            if (isZip && !isReCreate) {
                data = readFileToMemory();
            } else {
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                book.write(baos);
                data = baos.toByteArray();
            }
        } finally {
            deleteTempFile();
        }
        try (ByteArrayInputStream in = new ByteArrayInputStream(data)) {
            return uploader.upload(in, exportFileName);
        }
    }

    @Override
    public InputStream getInputStream(boolean excelCreated) {
        if (!excelCreated) {
            throw new RuntimeException("Excel has not been created yet; call createExcel() first");
        }
        if (isZip && !isReCreate) {
            try {
                byte[] data = readFileToMemory();
                return new ByteArrayInputStream(data);
            } catch (IOException e) {
                throw new RuntimeException("Failed to read Excel temporary file", e);
            } finally {
                deleteTempFile();
            }
        } else {
            try {
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                book.write(baos);
                return new ByteArrayInputStream(baos.toByteArray());
            } catch (IOException e) {
                throw new RuntimeException("Failed to serialize workbook to output stream", e);
            }
        }
    }

    @Override
    public void exportLocal(String filepath) {
        if (!filepath.endsWith(currentExcelType)) {
            filepath += "." + currentExcelType;
        }
        try (FileOutputStream fos = new FileOutputStream(filepath)) {
            if (isZip && !isReCreate) {
                try (FileInputStream fis = new FileInputStream(tempWorkFile)) {
                    fis.transferTo(fos);
                }
            } else {
                book.write(fos);
            }
            fos.flush();
            logger.debug("Export successful");
        } catch (IOException e) {
            logger.error("Export failed", e);
            throw new RuntimeException("Failed to export Excel to local file: " + filepath, e);
        } finally {
            deleteTempFile();
        }
    }

    private byte[] readFileToMemory() throws IOException {
        try (FileInputStream fis = new FileInputStream(tempWorkFile);
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[8192];
            int len;
            while ((len = fis.read(buffer)) != -1) {
                baos.write(buffer, 0, len);
            }
            return baos.toByteArray();
        }
    }

    private void deleteTempFile() {
        if (tempWorkFile != null) {
            tempWorkFile.delete();
            tempWorkFile = null;
        }
    }
}
