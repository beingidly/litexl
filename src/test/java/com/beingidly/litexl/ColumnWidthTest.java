package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class ColumnWidthTest {

    @Test
    void auto_isAutoWidth() {
        ColumnWidth auto = ColumnWidth.auto();
        assertTrue(auto.isAuto());
    }

    @Test
    void auto_getValueThrows() {
        ColumnWidth auto = ColumnWidth.auto();
        assertThrows(IllegalStateException.class, auto::getValue);
    }

    @Test
    void of_positiveValue() {
        ColumnWidth width = ColumnWidth.of(15.5);
        assertFalse(width.isAuto());
        assertEquals(15.5, width.getValue());
    }

    @Test
    void of_zeroThrows() {
        assertThrows(IllegalArgumentException.class, () -> ColumnWidth.of(0));
    }

    @Test
    void of_negativeThrows() {
        assertThrows(IllegalArgumentException.class, () -> ColumnWidth.of(-5));
    }

    @Test
    void toString_autoWidth() {
        assertEquals("ColumnWidth[auto]", ColumnWidth.auto().toString());
    }

    @Test
    void toString_explicitWidth() {
        assertEquals("ColumnWidth[15.5]", ColumnWidth.of(15.5).toString());
    }

    @Test
    void auto_isSingleton() {
        assertSame(ColumnWidth.auto(), ColumnWidth.auto());
    }
}
