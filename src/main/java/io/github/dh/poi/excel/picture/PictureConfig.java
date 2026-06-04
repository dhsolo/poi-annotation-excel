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
package io.github.dh.poi.excel.picture;

/**
 * Immutable value object that holds all picture-processing settings for an Excel export.
 *
 * <p>A {@code PictureConfig} governs how the framework fetches, resizes, and embeds images
 * into the exported workbook.  Instances are created exclusively through the fluent
 * {@link Builder}, which provides sensible defaults so that only non-default settings need
 * to be specified:
 * <ul>
 *   <li>Anchor type: {@link AnchorType#MOVE_AND_RESIZE}</li>
 *   <li>Image read timeout: 2 000 ms</li>
 *   <li>Images separator: {@code ","}</li>
 *   <li>Image resize: disabled</li>
 * </ul>
 *
 * <p>Example:
 * <pre>{@code
 * PictureConfig config = PictureConfig.builder()
 *         .anchorType(AnchorType.DONT_MOVE_AND_RESIZE)
 *         .imageReadTimeout(5000)
 *         .needImageResize(true)
 *         .resizeWidth(100)
 *         .resizeHeight(80)
 *         .build();
 * }</pre>
 *
 * @author dh
 * @since 1.0
 */
public class PictureConfig {

    /** Controls how an embedded image moves and resizes relative to its anchor cells. */
    private final AnchorType anchorType;

    /** Maximum time in milliseconds to wait when downloading a remote image URL. */
    private final int imageReadTimeout;

    /** Delimiter used to split a cell value that contains multiple image URLs. */
    private final String imagesSeparator;

    /** Whether downloaded images should be scaled before embedding. */
    private final boolean needImageResize;

    /** Target width in pixels when {@link #needImageResize} is {@code true}. */
    private final int resizeWidth;

    /** Target height in pixels when {@link #needImageResize} is {@code true}. */
    private final int resizeHeight;

    private PictureConfig(Builder builder) {
        this.anchorType = builder.anchorType;
        this.imageReadTimeout = builder.imageReadTimeout;
        this.imagesSeparator = builder.imagesSeparator;
        this.needImageResize = builder.needImageResize;
        this.resizeWidth = builder.resizeWidth;
        this.resizeHeight = builder.resizeHeight;
    }

    /**
     * Returns the anchor behaviour for embedded images.
     *
     * @return the {@link AnchorType}; never {@code null}
     */
    public AnchorType getAnchorType() { return anchorType; }

    /**
     * Returns the HTTP read timeout (in milliseconds) applied when downloading remote images.
     *
     * @return the timeout in ms; default is {@code 2000}
     */
    public int getImageReadTimeout() { return imageReadTimeout; }

    /**
     * Returns the delimiter string used to split a cell value that contains multiple image URLs.
     *
     * @return the separator; default is {@code ","}
     */
    public String getImagesSeparator() { return imagesSeparator; }

    /**
     * Returns whether images should be scaled to the configured dimensions before embedding.
     *
     * @return {@code true} if resize is enabled; default is {@code false}
     */
    public boolean isNeedImageResize() { return needImageResize; }

    /**
     * Returns the target pixel width for resized images.
     * Only meaningful when {@link #isNeedImageResize()} returns {@code true}.
     *
     * @return the resize width in pixels
     */
    public int getResizeWidth() { return resizeWidth; }

    /**
     * Returns the target pixel height for resized images.
     * Only meaningful when {@link #isNeedImageResize()} returns {@code true}.
     *
     * @return the resize height in pixels
     */
    public int getResizeHeight() { return resizeHeight; }

    /**
     * Creates a new {@link Builder} pre-populated with the default settings.
     *
     * @return a new builder instance
     */
    public static Builder builder() {
        return new Builder();
    }

    /**
     * Fluent builder for {@link PictureConfig}.
     *
     * <p>All setter methods return {@code this} to support method chaining.
     * Call {@link #build()} to obtain an immutable {@link PictureConfig} instance.
     */
    public static class Builder {

        private AnchorType anchorType = AnchorType.MOVE_AND_RESIZE;
        private int imageReadTimeout = 2000;
        private String imagesSeparator = ",";
        private boolean needImageResize = false;
        private int resizeWidth;
        private int resizeHeight;

        /**
         * Sets the anchor behaviour for embedded images.
         *
         * @param anchorType the desired anchor type; must not be {@code null}
         * @return this builder
         */
        public Builder anchorType(AnchorType anchorType) { this.anchorType = anchorType; return this; }

        /**
         * Sets the HTTP read timeout applied when downloading remote image URLs.
         *
         * @param timeout the timeout in milliseconds; must be &gt; 0
         * @return this builder
         */
        public Builder imageReadTimeout(int timeout) { this.imageReadTimeout = timeout; return this; }

        /**
         * Sets the delimiter used to split a cell value containing multiple image URLs.
         *
         * @param separator the separator string; must not be {@code null} or empty
         * @return this builder
         */
        public Builder imagesSeparator(String separator) { this.imagesSeparator = separator; return this; }

        /**
         * Enables or disables image resizing before the image is embedded.
         *
         * @param need {@code true} to resize images to the configured width and height
         * @return this builder
         */
        public Builder needImageResize(boolean need) { this.needImageResize = need; return this; }

        /**
         * Sets the target width (in pixels) when image resizing is enabled.
         *
         * @param width the desired width in pixels; must be &gt; 0 when resizing is enabled
         * @return this builder
         */
        public Builder resizeWidth(int width) { this.resizeWidth = width; return this; }

        /**
         * Sets the target height (in pixels) when image resizing is enabled.
         *
         * @param height the desired height in pixels; must be &gt; 0 when resizing is enabled
         * @return this builder
         */
        public Builder resizeHeight(int height) { this.resizeHeight = height; return this; }

        /**
         * Constructs and returns an immutable {@link PictureConfig} from the current builder state.
         *
         * @return a new {@code PictureConfig}; never {@code null}
         */
        public PictureConfig build() {
            return new PictureConfig(this);
        }
    }
}
