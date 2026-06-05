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
package io.github.dhsolo.poi.excel.export;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Strategy interface for exporting Excel workbooks to various targets.
 *
 * <p>Implemented by {@link DefaultExcelExporter}.
 *
 * @author dh
 * @since 1.0
 */
public interface ExcelExporter {

    /**
     * Sets the temporary workbook file produced when the export pipeline includes
     * picture embedding and ZIP re-packaging. Must be set before calling any export
     * method when {@link #setZip(boolean)} has been set to {@code true}.
     *
     * @param tempWorkFile the temporary file containing the packaged workbook
     */
    void setTempWorkFile(File tempWorkFile);

    /**
     * Signals that the workbook was re-packaged as a ZIP archive after picture embedding.
     * When {@code true} the export methods stream bytes from the temp file rather than
     * serialising the in-memory {@link Workbook}.
     *
     * @param zip {@code true} to read output from the temp file
     */
    void setZip(boolean zip);

    /**
     * Returns the in-memory POI {@link Workbook}.
     *
     * <p>When the workbook was ZIP-packaged ({@link #setZip(boolean)} {@code true}) and
     * has not yet been re-created from the temp file, this method reads the temp file,
     * deserialises it, and returns the resulting workbook so callers can make further
     * programmatic modifications before streaming.
     *
     * @return the current workbook instance
     */
    Workbook getWorkBook();

    /**
     * Writes the workbook content to the provided output stream and flushes it.
     * The temporary file (if any) is deleted after a successful write.
     *
     * @param outputStream   the target stream; the caller is responsible for closing it
     * @param exportFileName suggested file name used by some implementations for
     *                       Content-Disposition headers; may be ignored
     * @throws IOException if writing fails
     */
    void export(OutputStream outputStream, String exportFileName) throws IOException;

    /**
     * Serialises the workbook to a byte array and passes it as an {@link InputStream}
     * to the provided {@link ExcelUploader}, then returns the uploader's result.
     * The temporary file (if any) is deleted after the bytes have been read.
     *
     * @param uploader       custom storage backend (S3, MinIO, OSS, etc.)
     * @param exportFileName suggested file name forwarded to the uploader
     * @param <R>            the type returned by the uploader
     * @return the value returned by {@link ExcelUploader#upload}
     * @throws IOException if serialisation or the upload fails
     */
    <R> R upload(ExcelUploader<R> uploader, String exportFileName) throws IOException;

    /**
     * Returns the workbook content as an {@link InputStream} for downstream consumption
     * (e.g., attaching to an HTTP response or message payload).
     * The temporary file (if any) is deleted after the bytes have been read.
     *
     * @param excelCreated {@code true} if the workbook has been fully built; passing
     *                     {@code false} throws a {@link RuntimeException} immediately
     * @return an in-memory input stream backed by the serialised workbook bytes
     * @throws RuntimeException if {@code excelCreated} is {@code false} or if reading
     *                          the temp file fails
     */
    InputStream getInputStream(boolean excelCreated);

    /**
     * Writes the workbook to a local file at the given path. The file extension is
     * appended automatically if not already present.
     *
     * @param filepath absolute or relative path of the output file (extension optional)
     */
    void exportLocal(String filepath);
}
