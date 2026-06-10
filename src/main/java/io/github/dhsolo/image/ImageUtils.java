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
package io.github.dhsolo.image;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Image utility methods used by the Excel framework when embedding or processing
 * images inside Excel workbooks.
 *
 * <p>The class provides:
 * <ul>
 *   <li>URL encoding of Chinese characters so that image URLs containing CJK
 *       codepoints remain valid for HTTP requests.</li>
 *   <li>High-quality image resizing via Java2D, used before inserting pictures
 *       into cells to honour column/row dimension constraints.</li>
 * </ul>
 *
 * @author dhsolo
 * @since 1.0
 */
public class ImageUtils {

    private static final String DATA = "data";
    private static final String HTTP = "http";

    /**
     * Encodes any Chinese (CJK Unified Ideographs) characters found in a URL string
     * using UTF-8 percent-encoding, leaving all other characters unchanged.
     *
     * <p>This is necessary because image URLs stored in Excel data may contain raw
     * Chinese characters that are rejected by standard HTTP clients expecting
     * properly percent-encoded URIs.
     *
     * @param url the raw URL string, potentially containing CJK characters;
     *            must not be {@code null}
     * @return a new URL string in which every run of CJK characters has been
     *         replaced by its UTF-8 percent-encoded equivalent
     * @throws UnsupportedEncodingException if the UTF-8 charset is not available
     *                                      (practically never thrown on any JVM)
     */
    public static String urlEncoder(String url) throws UnsupportedEncodingException {
        Pattern p = Pattern.compile("[一-龥]+");
        Matcher matcher = p.matcher(url);
        StringBuffer b = new StringBuffer();
        while (matcher.find()) {
            matcher.appendReplacement(b, java.net.URLEncoder.encode(matcher.group(0), "utf-8"));
        }
        matcher.appendTail(b);
        return b.toString();
    }

    /**
     * Resizes a {@link BufferedImage} to the specified dimensions using bilinear
     * interpolation and anti-aliasing for high visual quality.
     *
     * <p>The resulting image has the same colour model type as the original.
     * The method is used by the framework to scale images before embedding them
     * in Excel cells so that they respect the column width and row height specified
     * in the export configuration.
     *
     * @param originalImage the source image to resize; must not be {@code null}
     * @param scaledWidth   the desired width of the output image in pixels;
     *                      must be greater than zero
     * @param scaledHeight  the desired height of the output image in pixels;
     *                      must be greater than zero
     * @return a new {@link BufferedImage} with the requested dimensions containing
     *         the rescaled content of {@code originalImage}
     * @throws IOException if an error occurs while creating the graphics context
     *                     (uncommon in practice)
     */
    public static BufferedImage resizeImage(BufferedImage originalImage, int scaledWidth, int scaledHeight) throws IOException {
        // ImageIO frequently decodes into TYPE_CUSTOM (0), which the BufferedImage constructor
        // rejects ("Unknown image type 0"); fall back to a standard type preserving alpha.
        int type = originalImage.getType();
        if (type == BufferedImage.TYPE_CUSTOM) {
            type = originalImage.getColorModel().hasAlpha()
                    ? BufferedImage.TYPE_INT_ARGB : BufferedImage.TYPE_INT_RGB;
        }
        BufferedImage resizedImage = new BufferedImage(scaledWidth, scaledHeight, type);
        Graphics2D g = resizedImage.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.drawImage(originalImage, 0, 0, scaledWidth, scaledHeight, null);
        g.dispose();
        return resizedImage;
    }
}
