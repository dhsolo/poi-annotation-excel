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

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.net.URL;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatCode;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class ImageDownloadPolicyTest {

    @AfterEach
    void reset() {
        ImageDownloadPolicy.setBlockPrivateNetworks(false);
    }

    @Test
    void allowsEverythingWhenDisabled() throws Exception {
        ImageDownloadPolicy.setBlockPrivateNetworks(false);
        assertThatCode(() -> ImageDownloadPolicy.assertAllowed(new URL("http://127.0.0.1/a.png")))
                .doesNotThrowAnyException();
    }

    @Test
    void blocksLoopbackAndPrivateWhenEnabled() {
        ImageDownloadPolicy.setBlockPrivateNetworks(true);
        assertThat(ImageDownloadPolicy.isBlockPrivateNetworks()).isTrue();
        assertThatThrownBy(() -> ImageDownloadPolicy.assertAllowed(new URL("http://127.0.0.1/a.png")))
                .isInstanceOf(IOException.class);
        assertThatThrownBy(() -> ImageDownloadPolicy.assertAllowed(new URL("http://10.1.2.3/a.png")))
                .isInstanceOf(IOException.class);
        assertThatThrownBy(() -> ImageDownloadPolicy.assertAllowed(new URL("http://192.168.0.5/a.png")))
                .isInstanceOf(IOException.class);
    }

    @Test
    void allowsPublicLiteralWhenEnabled() {
        ImageDownloadPolicy.setBlockPrivateNetworks(true);
        // 8.8.8.8 is a public literal IP — no DNS lookup, not private.
        assertThatCode(() -> ImageDownloadPolicy.assertAllowed(new URL("http://8.8.8.8/a.png")))
                .doesNotThrowAnyException();
    }
}
