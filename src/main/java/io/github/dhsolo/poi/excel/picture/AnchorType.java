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

/**
 * Controls how an embedded image anchor responds when the underlying cells are
 * moved or resized by the user or by programmatic column/row adjustments.
 *
 * <p>The integer values map directly to the POI {@code ClientAnchor} type constants
 * ({@code MOVE_AND_RESIZE = 0}, {@code MOVE_DONT_RESIZE = 2},
 * {@code DONT_MOVE_AND_RESIZE = 3}) so that the framework can pass them through to
 * the underlying POI API without additional translation.
 *
 * @author dh
 * @since 1.0
 * @see org.apache.poi.ss.usermodel.ClientAnchor.AnchorType
 */
public enum AnchorType {

    /**
     * The image moves and resizes together with the cells it is anchored to.
     * This is the default behaviour and is suitable for most use-cases where
     * images should track data layout changes.
     * Corresponds to POI constant {@code 0}.
     */
    MOVE_AND_RESIZE(0),

    /**
     * The image moves with the cells when rows or columns are inserted or deleted,
     * but its dimensions remain fixed regardless of cell size changes.
     * Corresponds to POI constant {@code 2}.
     */
    MOVE_DONT_RESIZE(2),

    /**
     * The image is absolutely positioned on the sheet; it neither moves nor resizes
     * when cells around it change. Useful for decorative elements that must stay at
     * a fixed position on the printable area.
     * Corresponds to POI constant {@code 3}.
     */
    DONT_MOVE_AND_RESIZE(3);

    /** The POI {@code ClientAnchor} integer constant for this anchor behaviour. */
    private final int value;

    AnchorType(int value) {
        this.value = value;
    }

    /**
     * Returns the POI {@code ClientAnchor} integer constant for this anchor type.
     *
     * @return the integer constant (0, 2, or 3)
     */
    public int getValue() {
        return this.value;
    }

    /**
     * Resolves an {@code AnchorType} from a POI anchor-type integer constant.
     * If the value does not match any constant, {@link #MOVE_AND_RESIZE} is returned
     * as the safe default.
     *
     * @param value the POI {@code ClientAnchor} integer constant to look up
     * @return the matching {@code AnchorType}, or {@link #MOVE_AND_RESIZE} if unknown
     */
    public static AnchorType fromValue(int value) {
        for (AnchorType type : values()) {
            if (type.value == value) {
                return type;
            }
        }
        return MOVE_AND_RESIZE;
    }
}
