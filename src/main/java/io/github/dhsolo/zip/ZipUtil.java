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
package io.github.dhsolo.zip;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Enumeration;
import java.util.HashSet;
import java.util.Set;
import java.util.zip.CRC32;
import java.util.zip.ZipEntry;

/**
 * ZIP file utilities based on JDK NIO ZipFileSystem.
 *
 * <p>Provides direct file injection into existing ZIP archives without
 * full extraction and re-compression.
 *
 * @author dhsolo
 * @since 1.0
 */
public class ZipUtil {

    private static final Logger logger = LoggerFactory.getLogger(ZipUtil.class);

    /**
     * Injects files from a source directory into an existing ZIP archive.
     *
     * <p>Files are placed preserving their relative path from {@code sourceDir}.
     * Existing entries with the same name are overwritten.
     *
     * @param zipPath   path to the existing ZIP file
     * @param sourceDir directory whose contents are injected
     * @throws IOException if an I/O error occurs
     */
    public static void injectDirectory(Path zipPath, Path sourceDir) throws IOException {
        try (FileSystem fs = FileSystems.newFileSystem(zipPath, (ClassLoader) null)) {
            Files.walkFileTree(sourceDir, new SimpleFileVisitor<Path>() {
                @Override
                public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
                    Path relative = sourceDir.relativize(dir);
                    if (!relative.toString().isEmpty()) {
                        Files.createDirectories(fs.getPath(relative.toString().replace('\\', '/')));
                    }
                    return FileVisitResult.CONTINUE;
                }

                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    Path relative = sourceDir.relativize(file);
                    Path dest = fs.getPath(relative.toString().replace('\\', '/'));
                    Files.createDirectories(dest.getParent());
                    Files.copy(file, dest, StandardCopyOption.REPLACE_EXISTING);
                    return FileVisitResult.CONTINUE;
                }
            });
        }
        logger.debug("inject directory into zip success: {}", zipPath);
    }

    /**
     * Injects specific files into an existing ZIP archive under a given prefix path by
     * rebuilding the archive once.
     *
     * <p>The injected files are images (JPEG/PNG/GIF/BMP) that are already compressed, so they
     * are added with the {@code STORED} method — re-deflating them only burns CPU (single
     * threaded) for negligible size gain. The injected entries carry their CRC and size in the
     * local header (no data descriptor), so they remain readable by streaming ZIP readers such
     * as POI's {@code XSSFWorkbook(InputStream)}.
     *
     * <p>All other entries (the workbook XML) are copied verbatim via
     * {@link ZipArchiveOutputStream#addRawArchiveEntry} — their already-compressed bytes are
     * transferred as-is, so the XML is never re-deflated either. Any existing entry whose name
     * collides with an injected file (e.g. an empty placeholder image part) is dropped in
     * favour of the injected one.
     *
     * @param zipPath   path to the existing ZIP file (rewritten in place)
     * @param files     files to inject
     * @param zipPrefix prefix path inside the ZIP (e.g. "xl/media")
     * @throws IOException if an I/O error occurs
     */
    public static void injectFiles(Path zipPath, File[] files, String zipPrefix) throws IOException {
        if (files == null || files.length == 0) return;

        Set<String> injected = new HashSet<>();
        for (File file : files) {
            injected.add(zipPrefix + "/" + file.getName());
        }

        Path rebuilt = Files.createTempFile(zipPath.getParent(), "rebuild", ".zip");
        try (ZipFile source = new ZipFile(zipPath.toFile());
             ZipArchiveOutputStream out = new ZipArchiveOutputStream(rebuilt.toFile())) {

            // Copy every existing entry verbatim (no re-compression), skipping placeholders
            // that the injected files replace.
            Enumeration<ZipArchiveEntry> entries = source.getEntriesInPhysicalOrder();
            while (entries.hasMoreElements()) {
                ZipArchiveEntry entry = entries.nextElement();
                if (injected.contains(entry.getName())) continue;
                out.addRawArchiveEntry(entry, source.getRawInputStream(entry));
            }

            // Add the images as STORED entries with header CRC/size.
            for (File file : files) {
                ZipArchiveEntry entry = new ZipArchiveEntry(zipPrefix + "/" + file.getName());
                entry.setMethod(ZipEntry.STORED);
                entry.setSize(file.length());
                entry.setCompressedSize(file.length());
                entry.setCrc(crc32(file));
                out.putArchiveEntry(entry);
                Files.copy(file.toPath(), out);
                out.closeArchiveEntry();
            }
        } catch (IOException e) {
            Files.deleteIfExists(rebuilt);
            throw e;
        }

        Files.move(rebuilt, zipPath, StandardCopyOption.REPLACE_EXISTING);
        logger.debug("inject {} files into zip success (stored, raw-copied): {}", files.length, zipPath);
    }

    private static long crc32(File file) throws IOException {
        CRC32 crc = new CRC32();
        byte[] buf = new byte[8192];
        try (InputStream in = new BufferedInputStream(new FileInputStream(file))) {
            int read;
            while ((read = in.read(buf)) != -1) {
                crc.update(buf, 0, read);
            }
        }
        return crc.getValue();
    }
}
