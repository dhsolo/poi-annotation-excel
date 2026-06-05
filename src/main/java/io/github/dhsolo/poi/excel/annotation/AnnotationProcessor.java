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
package io.github.dhsolo.poi.excel.annotation;

/**
 * Strategy interface for processing Excel annotations on model classes.
 *
 * <p>Extracts configuration from annotation-based Excel model definitions.
 * Implemented by {@link DefaultAnnotationProcessor}.
 *
 * @author dh
 * @since 1.0
 */
public interface AnnotationProcessor {

    /**
     * Processes the annotated model object and returns its complete Excel configuration.
     *
     * <p>The returned {@link ExcelAnnotationProperty} carries all
     * metadata extracted from the model's annotations: sheet info, header array, column
     * models, merge rules, column widths, cascade-validation models, title, custom rows,
     * and any nested complex sheet processors.
     *
     * @return a fully populated {@link ExcelAnnotationProperty};
     *         never {@code null}
     */
    ExcelAnnotationProperty getExcelAnnotationProperty();
}
