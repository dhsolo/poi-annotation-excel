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

import java.io.IOException;
import java.net.InetAddress;
import java.net.URL;
import java.net.UnknownHostException;

/**
 * Opt-in SSRF guard for remote image downloads.
 *
 * <p>Image URLs come from the data being exported and are therefore untrusted. By default this
 * library downloads whatever http/https URL it is given (preserving the long-standing behaviour
 * of fetching images from internal hosts when callers intend to). When the data source is not
 * fully trusted, enable {@link #setBlockPrivateNetworks(boolean) private-network blocking} so
 * that URLs resolving to loopback, link-local, site-local, or other private/reserved addresses
 * are rejected before any connection is made.
 *
 * <p>This is a process-wide policy; set it once at startup.
 *
 * @author dh
 * @since 1.0
 */
public final class ImageDownloadPolicy {

    private static volatile boolean blockPrivateNetworks = false;

    private ImageDownloadPolicy() {}

    /**
     * Enables or disables blocking of image URLs that resolve to private/reserved networks.
     * Defaults to {@code false} (no blocking) to preserve backward-compatible behaviour.
     *
     * @param block {@code true} to reject private/loopback/link-local targets
     */
    public static void setBlockPrivateNetworks(boolean block) {
        blockPrivateNetworks = block;
    }

    /** @return whether private-network blocking is currently enabled. */
    public static boolean isBlockPrivateNetworks() {
        return blockPrivateNetworks;
    }

    /**
     * Validates that the given URL is allowed to be fetched under the current policy.
     *
     * @param url the resolved image URL
     * @throws IOException if blocking is enabled and the host resolves to a private/reserved
     *                     address, or the host cannot be resolved
     */
    public static void assertAllowed(URL url) throws IOException {
        if (!blockPrivateNetworks) {
            return;
        }
        String host = url.getHost();
        if (host == null || host.isEmpty()) {
            throw new IOException("Blocked image URL with no host: " + url);
        }
        InetAddress[] addresses;
        try {
            addresses = InetAddress.getAllByName(host);
        } catch (UnknownHostException e) {
            throw new IOException("Cannot resolve image host: " + host, e);
        }
        for (InetAddress addr : addresses) {
            if (isPrivate(addr)) {
                throw new IOException("Blocked image URL targeting a private/reserved address: "
                        + host + " -> " + addr.getHostAddress());
            }
        }
    }

    private static boolean isPrivate(InetAddress addr) {
        return addr.isLoopbackAddress()
                || addr.isAnyLocalAddress()
                || addr.isLinkLocalAddress()
                || addr.isSiteLocalAddress()
                || addr.isMulticastAddress();
    }
}
