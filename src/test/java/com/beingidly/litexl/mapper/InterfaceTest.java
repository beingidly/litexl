package com.beingidly.litexl.mapper;

import com.beingidly.litexl.CellValue;
import com.beingidly.litexl.style.Style;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class InterfaceTest {

    static class IntConverter implements LitexlConverter<Integer> {
        @Override
        public Integer fromCell(CellValue value) {
            return switch (value) {
                case CellValue.Number n -> (int) n.value();
                default -> null;
            };
        }

        @Override
        public CellValue toCell(Integer value) {
            return value == null ? new CellValue.Empty() : new CellValue.Number(value);
        }
    }

    static class BoldStyle implements LitexlStyleProvider {
        @Override
        public Style provide() {
            return Style.builder().bold(true).build();
        }
    }

    @Test
    void converterFromCell() {
        var converter = new IntConverter();
        assertEquals(42, converter.fromCell(new CellValue.Number(42.0)));
        assertNull(converter.fromCell(new CellValue.Empty()));
    }

    @Test
    void converterToCell() {
        var converter = new IntConverter();
        assertEquals(new CellValue.Number(42), converter.toCell(42));
        assertEquals(new CellValue.Empty(), converter.toCell(null));
    }

    @Test
    void styleProvider() {
        var provider = new BoldStyle();
        var style = provider.provide();
        assertTrue(style.font().bold());
    }

    @Test
    void noneConverterThrows() {
        assertThrows(UnsupportedOperationException.class,
            () -> new LitexlConverter.None().fromCell(new CellValue.Empty()));
        assertThrows(UnsupportedOperationException.class,
            () -> new LitexlConverter.None().toCell(null));
    }
}
