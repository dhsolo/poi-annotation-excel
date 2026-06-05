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
package io.github.dhsolo.poi.excel;

import io.github.dhsolo.poi.excel.model.ComplexExcelModel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;

/**
 * Implements the Template Method pattern to orchestrate the fixed sequence of steps
 * required to produce an Excel workbook via {@link ExcelCreator}.
 *
 * <p>The pipeline is composed of nine ordered hook methods.  Each hook delegates to the
 * associated {@code ExcelCreator} operation, but subclasses may override any hook to inject
 * additional logic, skip a step, or replace the default behaviour entirely:
 * <ol>
 *   <li>{@link #initPictureDir()} — creates the temporary directory that stores downloaded
 *       images before they are embedded in the workbook.</li>
 *   <li>{@link #checkPicture()} — scans the data set and records which cells contain image
 *       URLs that must be fetched.</li>
 *   <li>{@link #headerAndTitle()} — writes the title row (if any) and the column header row.</li>
 *   <li>{@link #dataValidation()} — registers drop-down list validations for constrained columns.</li>
 *   <li>{@link #downloadPicture()} — fetches remote images and stores them in the temp directory.</li>
 *   <li>{@link #fillData()} — iterates over the data list and populates data rows cell by cell.</li>
 *   <li>{@link #mergeCells()} — applies any configured cell-merge regions.</li>
 *   <li>{@link #childSheets()} — recursively creates child sheets for complex multi-sheet models.</li>
 *   <li>{@link #postProcess()} — compresses the temp image directory if needed and wires up the
 *       ZIP artifact path on the exporter.</li>
 * </ol>
 *
 * <p>This class is package-private; callers interact with the framework through {@link ExcelCreator}.
 *
 * @author dh
 * @since 1.0
 */
class ExcelCreatePipeline {

    private static final Logger logger = LoggerFactory.getLogger(ExcelCreatePipeline.class);

    /** Ordered list of named pipeline steps registered in the constructor. */
    private final List<PipelineStep> steps = new ArrayList<>();

    /** The {@code ExcelCreator} instance whose operations back each pipeline hook. */
    private final ExcelCreator creator;

    /**
     * Constructs a pipeline bound to the given {@link ExcelCreator} and registers all
     * standard steps in the correct execution order.
     *
     * @param creator the workbook creator that provides the concrete operations for each step
     */
    ExcelCreatePipeline(ExcelCreator creator) {
        this.creator = creator;
        // Register standard pipeline steps
        steps.add(new PipelineStep("initPictureDir", this::initPictureDir));
        steps.add(new PipelineStep("checkPicture", this::checkPicture));
        steps.add(new PipelineStep("headerAndTitle", this::headerAndTitle));
        steps.add(new PipelineStep("dataValidation", this::dataValidation));
        steps.add(new PipelineStep("downloadPicture", this::downloadPicture));
        steps.add(new PipelineStep("fillData", this::fillData));
        steps.add(new PipelineStep("mergeCells", this::mergeCells));
        steps.add(new PipelineStep("childSheets", this::childSheets));
        steps.add(new PipelineStep("postProcess", this::postProcess));
    }

    /**
     * Runs all registered pipeline steps in order — this is the core Template Method entry point.
     *
     * <p>On completion, {@link ExcelCreator#setExcelCreated(boolean)} is set to {@code true}
     * and the total elapsed time is logged at {@code DEBUG} level.
     */
    void execute() {
        logger.debug("create excel start");
        long start = System.currentTimeMillis();

        for (PipelineStep step : steps) {
            step.run();
        }

        long duration = System.currentTimeMillis() - start;
        double time = new BigDecimal(duration).divide(new BigDecimal(1000), 8, RoundingMode.DOWN).doubleValue();
        logger.debug("create excel finished , spend time:{}", time);
        creator.setExcelCreated(true);
    }

    // ===== Overridable hook methods =====

    /**
     * Step 1 — Creates the temporary directory used to store downloaded image files.
     * Skipped when the current sheet is a child sheet of a complex workbook (the parent
     * already owns the temp directory).
     */
    protected void initPictureDir() {
        if (!creator.isChildComplex()) {
            creator.getPictureHandler().createTempleFileDir();
        }
    }

    /**
     * Step 2 — Scans the data set to identify cells that hold image URLs and prepares
     * the internal picture metadata needed by later steps.
     */
    protected void checkPicture() {
        creator.preparePictureData();
    }

    /**
     * Step 3 — Writes the optional title row and the mandatory column header row to the sheet.
     */
    protected void headerAndTitle() {
        creator.writeHeaderAndTitle(true);
    }

    /**
     * Step 4 — Registers drop-down list validations for columns that have constrained
     * value sets (e.g. yes/no, status codes), and records the current row offset so that
     * validation ranges start at the first data row.
     */
    protected void dataValidation() {
        creator.getDataValidator().checkListBox(
                creator.getColumnMappingInfo(), creator.isNeedOrderNum(), creator.getRowNum());
        creator.getDataValidator().setRowNum(creator.getRowNum());
        creator.setCurrentListNum(creator.getDataValidator().getCurrentListNum());
    }

    /**
     * Step 5 — Downloads remote images identified in step 2 and saves them to the
     * temporary directory created in step 1.
     */
    protected void downloadPicture() {
        creator.getPictureHandler().downLoadPicture();
    }

    /**
     * Step 6 — Iterates over the bound data list and writes each record to a row of
     * the sheet, resolving cell values and applying per-cell styles.
     */
    protected void fillData() {
        creator.populateData();
    }

    /**
     * Step 7 — Applies any configured cell-merge regions to the sheet, collapsing
     * adjacent cells with the same value into a single visual cell.
     */
    protected void mergeCells() {
        creator.mergeCells();
    }

    /**
     * Step 8 — Recursively generates child sheets for models that implement
     * {@link ComplexExcelModel}, each driven by its own sub-pipeline.
     */
    protected void childSheets() {
        creator.createChildSheets();
    }

    /**
     * Step 9 — Finalises the image artefact: compresses the temp image directory into
     * a ZIP archive if required, then propagates the ZIP flag and the temp-file path to
     * the exporter so it can deliver the correct output stream to the caller.
     * Skipped for child sheets because the parent sheet handles this step.
     */
    protected void postProcess() {
        if (!creator.isChildComplex()) {
            creator.getPictureHandler().decompressionPictureDirAndCompression();
            creator.getExporter().setZip(creator.getPictureHandler().isZip());
            creator.getExporter().setTempWorkFile(creator.getPictureHandler().getTempWorkFile());
        }
    }

    /**
     * Encapsulates a single named pipeline step, pairing a human-readable name used in
     * debug log messages with the {@link Runnable} action that performs the work.
     */
    private static class PipelineStep {

        /** Human-readable name for this step, used in debug log messages. */
        final String name;

        /** The action to execute when this step is run. */
        final Runnable action;

        /**
         * Constructs a pipeline step with the given name and action.
         *
         * @param name   a short descriptive label for logging
         * @param action the logic to execute for this step
         */
        PipelineStep(String name, Runnable action) {
            this.name = name;
            this.action = action;
        }

        /**
         * Logs the step name at {@code DEBUG} level and then executes the action.
         */
        void run() {
            logger.debug("pipeline step: {}", name);
            action.run();
        }
    }
}
