package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class ColumnIndexTest {

    @Test
    void none_returnsNonExistentIndex() {
        ColumnIndex none = ColumnIndex.none();
        assertFalse(none.exists());
    }

    @Test
    void none_getValueThrows() {
        ColumnIndex none = ColumnIndex.none();
        assertThrows(IllegalStateException.class, none::getValue);
    }

    @Test
    void of_validIndex() {
        ColumnIndex idx = ColumnIndex.of(5);
        assertTrue(idx.exists());
        assertEquals(5, idx.getValue());
    }

    @Test
    void of_zeroIsValid() {
        ColumnIndex idx = ColumnIndex.of(0);
        assertTrue(idx.exists());
        assertEquals(0, idx.getValue());
    }

    @Test
    void of_negativeThrows() {
        assertThrows(IllegalArgumentException.class, () -> ColumnIndex.of(-1));
    }

    @Test
    void toString_existingIndex() {
        assertEquals("ColumnIndex[5]", ColumnIndex.of(5).toString());
    }

    @Test
    void toString_noneIndex() {
        assertEquals("ColumnIndex[none]", ColumnIndex.none().toString());
    }
}
