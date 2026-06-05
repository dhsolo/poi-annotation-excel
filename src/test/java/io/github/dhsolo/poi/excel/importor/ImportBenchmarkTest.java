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
package io.github.dhsolo.poi.excel.importor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIfSystemProperty;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * Manual benchmark comparing large {@code .xlsx} import via the DOM-based POI workbook
 * (what {@link ExcelImportor} uses) against the SAX-based {@link StreamingExcelReader}.
 *
 * <p>It is <strong>disabled by default</strong> so it never runs in the normal build. Enable
 * it explicitly:
 *
 * <pre>{@code
 * mvn -o test -Dtest=ImportBenchmarkTest -Dexcel.benchmark=true
 * # optional knobs:
 * #   -Dexcel.benchmark.rows=200000   (default 100000)
 * #   -Dexcel.benchmark.cols=8        (default 6)
 * }</pre>
 *
 * <p>Memory figures are retained-heap deltas (after {@code System.gc()}): the DOM workbook is
 * held while measured, the streaming reader retains nothing. Numbers are indicative, not
 * exact, and depend on the JVM heap given to the forked test process.
 */
@EnabledIfSystemProperty(named = "excel.benchmark", matches = "true")
class ImportBenchmarkTest {

    @Test
    void domVsStreaming(@TempDir File tmp) throws Exception {
        int rows = Integer.getInteger("excel.benchmark.rows", 100_000);
        int cols = Integer.getInteger("excel.benchmark.cols", 6);

        File file = new File(tmp, "bench.xlsx");
        generate(file, rows, cols);
        long fileKb = file.length() / 1024;

        System.out.printf("%n=== Import benchmark: %,d rows x %d cols (%,d KB on disk) ===%n",
                rows, cols, fileKb);

        // SAX first, while the heap is clean — it retains nothing, so its baseline is not
        // polluted by a previously loaded DOM workbook.
        long saxBase = usedMem();
        long saxStart = System.nanoTime();
        long[] saxCells = {0};
        try (FileInputStream in = new FileInputStream(file)) {
            StreamingExcelReader.read(in, 0, (rowIndex, cells) -> {
                for (String s : cells) saxCells[0] += s.length();
            });
        }
        long saxMem = usedMem() - saxBase;
        long saxMs = (System.nanoTime() - saxStart) / 1_000_000;

        // DOM (POI WorkbookFactory, as used by ExcelImportor) — measured while the workbook
        // is still referenced so the figure reflects retained heap.
        long domBase = usedMem();
        long domStart = System.nanoTime();
        long domCells;
        long domMem;
        try (Workbook wb = WorkbookFactory.create(file)) {
            long sum = 0;
            Sheet sheet = wb.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    sum += text(cell).length();
                }
            }
            domCells = sum;
            domMem = usedMem() - domBase; // workbook still referenced here
        }
        long domMs = (System.nanoTime() - domStart) / 1_000_000;

        System.out.printf("%-22s %10s %14s%n", "approach", "time(ms)", "retained(MB)");
        System.out.printf("%-22s %10d %14.1f%n", "DOM (WorkbookFactory)", domMs, domMem / 1048576.0);
        System.out.printf("%-22s %10d %14.1f%n", "SAX (StreamingReader)", saxMs, saxMem / 1048576.0);
        System.out.printf("checksum dom=%d sax=%d (should match)%n", domCells, saxCells[0]);
    }

    private static void generate(File file, int rows, int cols) throws Exception {
        try (SXSSFWorkbook wb = new SXSSFWorkbook(100); FileOutputStream fo = new FileOutputStream(file)) {
            Sheet sheet = wb.createSheet("bench");
            Row header = sheet.createRow(0);
            for (int c = 0; c < cols; c++) header.createCell(c).setCellValue("col" + c);
            for (int r = 1; r <= rows; r++) {
                Row row = sheet.createRow(r);
                for (int c = 0; c < cols; c++) {
                    row.createCell(c).setCellValue("v-" + r + "-" + c);
                }
            }
            wb.write(fo);
            wb.dispose();
        }
    }

    private static String text(Cell cell) {
        try {
            return switch (cell.getCellType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                default -> "";
            };
        } catch (Exception e) {
            return "";
        }
    }

    private static long usedMem() {
        Runtime rt = Runtime.getRuntime();
        for (int i = 0; i < 3; i++) {
            rt.gc();
            try { Thread.sleep(60); } catch (InterruptedException ignored) { Thread.currentThread().interrupt(); }
        }
        return rt.totalMemory() - rt.freeMemory();
    }
}
