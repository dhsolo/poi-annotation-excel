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

import com.sun.net.httpserver.HttpServer;
import io.github.dhsolo.poi.excel.ExcelUtil;
import io.github.dhsolo.poi.excel.annotation.ExcelColumn;
import io.github.dhsolo.poi.excel.annotation.ExcelData;
import io.github.dhsolo.poi.excel.annotation.ExcelImage;
import io.github.dhsolo.poi.excel.annotation.ExcelInfo;
import org.junit.jupiter.api.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.net.InetSocketAddress;
import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.CountDownLatch;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Concurrent-export stress/smoke test for the shared image-download pool lifecycle: several
 * threads repeatedly export image-bearing models, so the pool is constantly created (each export
 * increments {@code INSTANCE_COUNT} and grabs the pool) and retired (the last close drains it),
 * with download tasks submitted throughout. Exercises the create/close/shutdown cycling under
 * contention and asserts every export still completes with a non-empty workbook.
 *
 * <p>Note: this guards against gross breakage (deadlock, NPE on the null-pool recreation path,
 * rejected submissions) but does not <em>deterministically</em> reproduce the narrow shut-down-
 * under-active-export window the {@code ExcelCreator.close()} re-check fixes — that interleave is
 * timing-dependent and rare. The fix's correctness rests on the synchronization argument
 * documented at the close() re-check.
 */
class ConcurrentPictureExportTest {

    @Test
    void concurrentImageExportsAllSucceed() throws Exception {
        byte[] png = png();
        HttpServer server = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        server.createContext("/logo.png", exchange -> {
            exchange.getResponseHeaders().add("Content-Type", "image/png");
            exchange.sendResponseHeaders(200, png.length);
            try (OutputStream os = exchange.getResponseBody()) {
                os.write(png);
            }
        });
        server.start();
        String url = "http://127.0.0.1:" + server.getAddress().getPort() + "/logo.png";

        // Few threads with a small per-iteration jitter so the active-creator count repeatedly
        // crosses zero (one thread retiring the pool while another is just starting its export) —
        // the exact window the fix protects.
        int threads = 3, perThread = 40;
        CountDownLatch start = new CountDownLatch(1);
        CountDownLatch done = new CountDownLatch(threads);
        List<Throwable> failures = new CopyOnWriteArrayList<>();
        List<Integer> sizes = new CopyOnWriteArrayList<>();

        try {
            for (int t = 0; t < threads; t++) {
                final int seed = t;
                new Thread(() -> {
                    try {
                        start.await();
                        for (int i = 0; i < perThread; i++) {
                            byte[] bytes = ExcelUtil.toBytes(
                                    new ImgModel().setData(List.of(new Row("a", url), new Row("b", url))));
                            sizes.add(bytes.length);
                            if (((i + seed) & 1) == 0) Thread.sleep(1); // desync the create/close cycles
                        }
                    } catch (Throwable e) {
                        failures.add(e);
                    } finally {
                        done.countDown();
                    }
                }, "export-" + t).start();
            }
            start.countDown();
            done.await();
        } finally {
            server.stop(0);
        }

        assertThat(failures).as("every concurrent export should complete without error (e.g. no "
                + "RejectedExecutionException from a pool shut down underneath an active export)").isEmpty();
        assertThat(sizes).hasSize(threads * perThread).allSatisfy(s -> assertThat(s).isGreaterThan(0));
    }

    private static byte[] png() throws Exception {
        BufferedImage img = new BufferedImage(6, 6, BufferedImage.TYPE_INT_RGB);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ImageIO.write(img, "png", out);
        return out.toByteArray();
    }

    @ExcelInfo(sheetName = "p", isBigData = false)
    public static class ImgModel {
        @ExcelData
        private List<Row> data;
        @ExcelColumn(columnName = "N", index = 1)
        private String name;
        @ExcelColumn(columnName = "P", index = 2)
        @ExcelImage
        private String pic;

        public ImgModel setData(List<Row> data) {
            this.data = data;
            return this;
        }

        public List<Row> getData() {
            return data;
        }
    }

    public static class Row {
        private final String name;
        private final String pic;

        public Row(String name, String pic) {
            this.name = name;
            this.pic = pic;
        }

        public String getName() {
            return name;
        }

        public String getPic() {
            return pic;
        }
    }
}
