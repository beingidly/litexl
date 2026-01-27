package com.beingidly.litexl.mapper;

import com.beingidly.litexl.CellValue;
import org.jspecify.annotations.Nullable;

public interface LitexlConverter<T> {

    @Nullable
    T fromCell(CellValue value);

    CellValue toCell(@Nullable T value);

    final class None implements LitexlConverter<Object> {
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
