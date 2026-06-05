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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.awt.image.ColorModel;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilterInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.util.Arrays;
import java.util.Map;

import static io.github.dhsolo.image.ImageUtils.resizeImage;
import static io.github.dhsolo.image.ImageUtils.urlEncoder;

/**
 * Downloads a batch of images to the temporary directory, one file per image index.
 *
 * <p>Each image is stored as {@code image{index}.{ext}} where {@code ext} matches the
 * {@link PictureFormat} pre-registered for that index (derived from the URL), so the file can
 * later be injected straight into the workbook's {@code xl/media} folder.
 *
 * <p>To keep heap usage low and avoid lossy re-encoding, an image is copied byte-for-byte when
 * no resize is requested and its sniffed format equals the declared format; only the resize
 * path (or a format mismatch) decodes the image into a {@link BufferedImage}.
 */
public class ImageDownLoadTask implements Runnable {

    private static final Logger log = LoggerFactory.getLogger(ImageDownLoadTask.class);
    private static final int MAX_RETRY = 3;
    private static final int MAGIC_PEEK = 16;
    /** Hard cap on a single downloaded image to guard against oversized/hostile responses. */
    private static final long MAX_IMAGE_BYTES = 64L * 1024 * 1024;

    private final Map<Integer, String> pictureIndexMapping;
    private final String path;
    private final int imageReadTimeOut;
    private final boolean needResize;
    private final int resizeWith;
    private final int resizeHeight;
    private final int taskIndex;

    public ImageDownLoadTask(Map<Integer, String> pictureIndexMapping, String path, int imageReadTimeOut,
                             boolean needResize, int resizeWith, int resizeHeight, int taskIndex) {
        this.pictureIndexMapping = pictureIndexMapping;
        this.path = path;
        this.imageReadTimeOut = imageReadTimeOut;
        this.needResize = needResize;
        this.resizeWith = resizeWith;
        this.resizeHeight = resizeHeight;
        this.taskIndex = taskIndex;
    }

    @Override
    public void run() {
        pictureIndexMapping.forEach((index, url) -> {
            PictureFormat declared = PictureFormat.fromUrl(url);
            File targetFile = new File(path, "image" + index + "." + declared.fileExtension());
            File tempFile = new File(path, "image" + index + ".tmp");

            boolean stored = false;
            for (int attempt = 1; attempt <= MAX_RETRY && !stored; attempt++) {
                try {
                    storeImage(url, declared, tempFile);
                    stored = true;
                } catch (Exception e) {
                    log.warn("Image download failed, retrying ({}/{}): index={}, url={}", attempt, MAX_RETRY, index, url, e);
                    deleteQuietly(tempFile);
                }
            }

            if (stored) {
                if (!tempFile.renameTo(targetFile)) {
                    log.warn("Failed to rename temp file to target: {}", targetFile.getName());
                    deleteQuietly(tempFile);
                }
            } else {
                // Permanent failure: write a placeholder so the pre-registered relationship still
                // resolves and the workbook does not open as a corrupt file.
                log.error("Image download failed permanently, writing placeholder: index={}, url={}", index, url);
                try {
                    writePlaceholder(declared, targetFile);
                } catch (Exception ex) {
                    log.error("Failed to write placeholder image: index={}", index, ex);
                }
            }
        });
        log.debug("finish down picture task : {}", taskIndex);
    }

    /**
     * Downloads {@code url} and stores it into {@code tempFile} in the {@code declared} format.
     * Streams the raw bytes verbatim when no resize is needed and the sniffed format already
     * matches; otherwise decodes, optionally flattens transparency / resizes, and re-encodes.
     */
    private void storeImage(String url, PictureFormat declared, File tempFile) throws Exception {
        try (BufferedInputStream in = new BufferedInputStream(openStream(url))) {
            byte[] header = peek(in, MAGIC_PEEK);
            PictureFormat actual = PictureFormat.sniff(header);

            if (!needResize && actual == declared) {
                try (OutputStream os = new FileOutputStream(tempFile)) {
                    in.transferTo(os);
                    os.flush();
                }
                return;
            }

            BufferedImage img = ImageIO.read(in);
            if (img == null) {
                throw new IOException("Unable to decode image: " + url);
            }
            if (!declared.supportsAlpha()) {
                img = flattenToRgb(img);
            }
            if (needResize) {
                img = resizeImage(img, resizeWith, resizeHeight);
            }
            try (OutputStream os = new FileOutputStream(tempFile)) {
                if (!ImageIO.write(img, declared.imageIoName(), os)) {
                    throw new IOException("No ImageIO writer for format: " + declared.imageIoName());
                }
                os.flush();
            }
            img.flush();
        }
    }

    /**
     * Opens an input stream for the image. Remote URLs must use http/https, are CJK-encoded,
     * time-bounded, and rejected early if the declared content length exceeds the size cap.
     * Any other value is treated as a local file path. The returned stream is size-limited
     * to {@link #MAX_IMAGE_BYTES} to defend against oversized or hostile responses.
     */
    private InputStream openStream(String url) throws Exception {
        InputStream raw;
        if (url.startsWith("http")) {
            URL imgUrl = new URL(urlEncoder(url));
            String protocol = imgUrl.getProtocol();
            if (!"http".equalsIgnoreCase(protocol) && !"https".equalsIgnoreCase(protocol)) {
                throw new IOException("Unsupported image URL protocol: " + protocol);
            }
            ImageDownloadPolicy.assertAllowed(imgUrl);
            URLConnection conn = imgUrl.openConnection();
            // In secure mode, do not auto-follow redirects: a 30x to an internal host would
            // otherwise bypass the private-network check performed above.
            if (conn instanceof HttpURLConnection httpConn && ImageDownloadPolicy.isBlockPrivateNetworks()) {
                httpConn.setInstanceFollowRedirects(false);
            }
            conn.setConnectTimeout(imageReadTimeOut);
            conn.setReadTimeout(imageReadTimeOut);
            long declaredLength = conn.getContentLengthLong();
            if (declaredLength > MAX_IMAGE_BYTES) {
                throw new IOException("Image exceeds max size (" + declaredLength + " > " + MAX_IMAGE_BYTES + " bytes): " + url);
            }
            raw = conn.getInputStream();
        } else {
            raw = new FileInputStream(new File(url));
        }
        return new LimitedInputStream(raw, MAX_IMAGE_BYTES);
    }

    /** Wraps a stream so reads beyond {@code limit} bytes fail fast instead of buffering unbounded data. */
    private static final class LimitedInputStream extends FilterInputStream {
        private final long limit;
        private long count;

        LimitedInputStream(InputStream in, long limit) {
            super(in);
            this.limit = limit;
        }

        @Override
        public int read() throws IOException {
            int b = super.read();
            if (b >= 0 && ++count > limit) {
                throw new IOException("Image exceeds max size " + limit + " bytes");
            }
            return b;
        }

        @Override
        public int read(byte[] b, int off, int len) throws IOException {
            int n = super.read(b, off, len);
            if (n > 0) {
                count += n;
                if (count > limit) {
                    throw new IOException("Image exceeds max size " + limit + " bytes");
                }
            }
            return n;
        }
    }

    /** Reads up to {@code n} leading bytes without consuming the stream. */
    private static byte[] peek(BufferedInputStream in, int n) throws IOException {
        in.mark(n + 1);
        byte[] buf = new byte[n];
        int read = 0;
        while (read < n) {
            int r = in.read(buf, read, n - read);
            if (r < 0) break;
            read += r;
        }
        in.reset();
        return read == n ? buf : Arrays.copyOf(buf, read);
    }

    /** Flattens any alpha channel onto a white background so the image can be encoded as JPEG. */
    private BufferedImage flattenToRgb(BufferedImage img) {
        ColorModel cm = img.getColorModel();
        if (!cm.hasAlpha() && cm.getPixelSize() != 32) return img;
        BufferedImage rgb = new BufferedImage(img.getWidth(), img.getHeight(), BufferedImage.TYPE_INT_RGB);
        Graphics2D g = rgb.createGraphics();
        try {
            g.drawImage(img, 0, 0, Color.WHITE, null);
        } finally {
            g.dispose();
        }
        img.flush();
        return rgb;
    }

    /** Writes a minimal 1x1 placeholder in the declared format. */
    private void writePlaceholder(PictureFormat declared, File target) throws IOException {
        int type = declared.supportsAlpha() ? BufferedImage.TYPE_INT_ARGB : BufferedImage.TYPE_INT_RGB;
        BufferedImage img = new BufferedImage(1, 1, type);
        if (!declared.supportsAlpha()) {
            img.setRGB(0, 0, 0xFFFFFF);
        }
        try (OutputStream os = new FileOutputStream(target)) {
            ImageIO.write(img, declared.imageIoName(), os);
            os.flush();
        }
    }

    private static void deleteQuietly(File file) {
        if (file != null && file.exists() && !file.delete()) {
            log.debug("Could not delete temp file: {}", file.getName());
        }
    }
}
