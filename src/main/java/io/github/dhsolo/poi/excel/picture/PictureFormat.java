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

import org.apache.poi.xssf.usermodel.XSSFRelation;

/**
 * Describes an embeddable picture format and centralises every per-format detail the
 * framework needs: the file extension used inside the workbook's {@code xl/media} folder,
 * the {@link javax.imageio.ImageIO} writer name, the POI {@link XSSFRelation} that declares
 * the part's content type, the legacy {@code Workbook.PICTURE_TYPE_*} constant, and whether
 * the format carries an alpha (transparency) channel.
 *
 * <p>The {@link #fileExtension()} intentionally matches the extension POI assigns to the
 * corresponding {@link XSSFRelation} part name (e.g. {@code IMAGE_PNG} → {@code image#.png}).
 * Keeping the on-disk temp file name (written by {@code ImageDownLoadTask}) identical to the
 * pre-created relationship part name is what allows downloaded images to be injected straight
 * into the ZIP without POI re-processing them.
 *
 * @author dh
 * @since 1.0
 */
public enum PictureFormat {

    /** JPEG / JPG — no transparency; alpha is flattened onto white before encoding. */
    JPEG("jpeg", "jpg", XSSFRelation.IMAGE_JPEG, false),
    /** PNG — lossless, preserves transparency. */
    PNG("png", "png", XSSFRelation.IMAGE_PNG, true),
    /** GIF — preserves transparency (single frame when re-encoded). */
    GIF("gif", "gif", XSSFRelation.IMAGE_GIF, true),
    /** BMP / DIB — no transparency. */
    BMP("bmp", "bmp", XSSFRelation.IMAGE_BMP, false);

    private final String fileExtension;
    private final String imageIoName;
    private final XSSFRelation relation;
    private final boolean supportsAlpha;

    PictureFormat(String fileExtension, String imageIoName, XSSFRelation relation,
                  boolean supportsAlpha) {
        this.fileExtension = fileExtension;
        this.imageIoName = imageIoName;
        this.relation = relation;
        this.supportsAlpha = supportsAlpha;
    }

    /** @return the file extension (no dot) matching the POI relation part name. */
    public String fileExtension() {
        return fileExtension;
    }

    /** @return the {@link javax.imageio.ImageIO} informal format name used for encoding. */
    public String imageIoName() {
        return imageIoName;
    }

    /** @return the POI relationship that declares this format's content type. */
    public XSSFRelation relation() {
        return relation;
    }

    /** @return {@code true} when this format carries an alpha (transparency) channel. */
    public boolean supportsAlpha() {
        return supportsAlpha;
    }

    /**
     * Determines the intended format from a URL (or local path) by inspecting its file
     * extension. Query strings and fragments are ignored. Unknown or extension-less inputs
     * fall back to {@link #JPEG}.
     *
     * @param url the image URL or local file path; may be {@code null}
     * @return the matching {@code PictureFormat}, never {@code null}
     */
    public static PictureFormat fromUrl(String url) {
        if (url == null) return JPEG;
        String u = url.trim().toLowerCase();
        int cut = u.indexOf('?');
        if (cut >= 0) u = u.substring(0, cut);
        cut = u.indexOf('#');
        if (cut >= 0) u = u.substring(0, cut);
        int dot = u.lastIndexOf('.');
        int slash = Math.max(u.lastIndexOf('/'), u.lastIndexOf('\\'));
        if (dot < 0 || dot < slash) return JPEG;
        return switch (u.substring(dot + 1)) {
            case "png" -> PNG;
            case "gif" -> GIF;
            case "bmp", "dib" -> BMP;
            default -> JPEG; // jpg, jpeg, jpe, unknown
        };
    }

    /**
     * Detects the actual format of raw image bytes from their magic-number header.
     *
     * @param header the leading bytes of the image (at least 8 recommended); may be {@code null}
     * @return the detected {@code PictureFormat}, or {@code null} if the bytes match no known format
     */
    public static PictureFormat sniff(byte[] header) {
        if (header == null) return null;
        int n = header.length;
        if (n >= 3 && (header[0] & 0xFF) == 0xFF && (header[1] & 0xFF) == 0xD8 && (header[2] & 0xFF) == 0xFF) {
            return JPEG;
        }
        if (n >= 8 && (header[0] & 0xFF) == 0x89 && header[1] == 'P' && header[2] == 'N' && header[3] == 'G'
                && (header[4] & 0xFF) == 0x0D && (header[5] & 0xFF) == 0x0A
                && (header[6] & 0xFF) == 0x1A && (header[7] & 0xFF) == 0x0A) {
            return PNG;
        }
        if (n >= 3 && header[0] == 'G' && header[1] == 'I' && header[2] == 'F') {
            return GIF;
        }
        if (n >= 2 && header[0] == 'B' && header[1] == 'M') {
            return BMP;
        }
        return null;
    }
}
