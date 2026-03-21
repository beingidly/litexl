package com.beingidly.litexl.mapper;

import com.beingidly.litexl.CellValue;
import org.jspecify.annotations.Nullable;

/**
 * Converts between cell values and Java types.
 *
 * @param <T> the Java type this converter handles
 */
public interface LitexlConverter<T> {

    /**
     * Converts a cell value to the target Java type.
     *
     * @param value the cell value
     * @return the converted value, or null
     */
    @Nullable
    T fromCell(CellValue value);

    /**
     * Converts a Java value to a cell value.
     *
     * @param value the Java value, or null
     * @return the cell value
     */
    CellValue toCell(@Nullable T value);

    /** Placeholder converter indicating no custom conversion. */
    final class None implements LitexlConverter<Object> {
        /** Creates a None converter instance. */
        public None() {}

        @Override
        public @Nullable Object fromCell(CellValue value) {
            throw new UnsupportedOperationException("None converter cannot be used");
        }

        @Override
        public CellValue toCell(@Nullable Object value) {
            throw new UnsupportedOperationException("None converter cannot be used");
        }
    }
}
