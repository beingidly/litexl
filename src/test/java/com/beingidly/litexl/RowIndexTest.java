package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class RowIndexTest {

    @Test
    void none_returnsNonExistentIndex() {
        RowIndex none = RowIndex.none();
        assertFalse(none.exists());
    }

    @Test
    void none_getValueThrows() {
        RowIndex none = RowIndex.none();
        assertThrows(IllegalStateException.class, none::getValue);
    }

    @Test
    void of_validIndex() {
        RowIndex idx = RowIndex.of(10);
        assertTrue(idx.exists());
        assertEquals(10, idx.getValue());
    }

    @Test
    void of_zeroIsValid() {
        RowIndex idx = RowIndex.of(0);
        assertTrue(idx.exists());
        assertEquals(0, idx.getValue());
    }

    @Test
    void of_negativeThrows() {
        assertThrows(IllegalArgumentException.class, () -> RowIndex.of(-1));
    }

    @Test
    void toString_existingIndex() {
        assertEquals("RowIndex[10]", RowIndex.of(10).toString());
    }

    @Test
    void toString_noneIndex() {
        assertEquals("RowIndex[none]", RowIndex.none().toString());
    }
}
