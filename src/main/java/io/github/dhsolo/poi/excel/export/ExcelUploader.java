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

import java.io.IOException;
import java.io.InputStream;

/**
 * Strategy for uploading an exported Excel file to a remote storage backend.
 *
 * <p>Implement this interface to connect any storage system — S3, MinIO, OSS,
 * an internal file server, FTP, etc. — without changing the export pipeline.
 * The type parameter {@code R} is the result returned by the storage backend
 * (e.g. a URL {@code String}, a file ID {@code Long}, a response object, etc.).
 * Use {@code Void} when no result is needed.
 *
 * <p>Example — AWS S3 (returns ETag string):
 * <pre>
 * String etag = ExcelUtil.upload(
 *     (in, name) -> s3Client.putObject(b -> b.bucket(bucket).key(name),
 *                                      RequestBody.fromInputStream(in, in.available()))
 *                           .eTag(),
 *     "report.xlsx", builder);
 * </pre>
 *
 * <p>Example — internal file service (returns file URL):
 * <pre>
 * String url = ExcelUtil.upload(fileService::store, "report.xlsx", model);
 * </pre>
 *
 * <p>Example — no return value needed:
 * <pre>
 * ExcelUtil.&lt;Void&gt;upload((in, name) -> { ftp.upload(in, name); return null; }, "report.xlsx", model);
 * </pre>
 *
 * @param <R> the type of result returned after a successful upload
 * @author dhsolo
 * @since 1.0
 */
@FunctionalInterface
public interface ExcelUploader<R> {

    /**
     * Uploads the Excel content to a storage backend and returns a result.
     *
     * @param inputStream content to upload — the caller closes the stream after this method returns
     * @param fileName    suggested file name including extension (e.g. {@code "report.xlsx"})
     * @return upload result (e.g. URL, file ID, response object); may be {@code null}
     * @throws IOException if the upload fails
     */
    R upload(InputStream inputStream, String fileName) throws IOException;
}
