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

import org.junit.jupiter.api.Test;

import java.awt.Transparency;
import java.awt.color.ColorSpace;
import java.awt.image.BufferedImage;
import java.awt.image.ColorModel;
import java.awt.image.ComponentColorModel;
import java.awt.image.DataBuffer;
import java.awt.image.WritableRaster;

import static org.assertj.core.api.Assertions.assertThat;

class ImageUtilsTest {

    /**
     * Regression: ImageIO frequently decodes into TYPE_CUSTOM (0); building the resize target
     * with the original type crashed with "Unknown image type 0", so every resize of such an
     * image failed and fell back to a placeholder.
     */
    @Test
    void resizeHandlesTypeCustomImages() throws Exception {
        ColorModel cm = new ComponentColorModel(ColorSpace.getInstance(ColorSpace.CS_sRGB),
                true, false, Transparency.TRANSLUCENT, DataBuffer.TYPE_BYTE);
        WritableRaster raster = cm.createCompatibleWritableRaster(8, 8);
        BufferedImage custom = new BufferedImage(cm, raster, false, null);
        assertThat(custom.getType()).isEqualTo(BufferedImage.TYPE_CUSTOM);

        BufferedImage resized = ImageUtils.resizeImage(custom, 4, 4);
        assertThat(resized.getWidth()).isEqualTo(4);
        assertThat(resized.getHeight()).isEqualTo(4);
        assertThat(resized.getColorModel().hasAlpha()).isTrue();
    }

    @Test
    void resizeKeepsStandardTypes() throws Exception {
        BufferedImage rgb = new BufferedImage(8, 8, BufferedImage.TYPE_INT_RGB);
        BufferedImage resized = ImageUtils.resizeImage(rgb, 2, 2);
        assertThat(resized.getType()).isEqualTo(BufferedImage.TYPE_INT_RGB);
    }
}
