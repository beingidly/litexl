package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.*;

class CellRangeTest {

    @Test
    void createRange() {
        CellRange range = CellRange.of(0, 0, 5, 5);

        assertEquals(0, range.startRow());
        assertEquals(0, range.startCol());
        assertEquals(5, range.endRow());
        assertEquals(5, range.endCol());
    }

    @Test
    void createSingleCellRange() {
        CellRange range = CellRange.of(3, 2);

        assertEquals(3, range.startRow());
        assertEquals(2, range.startCol());
        assertEquals(3, range.endRow());
        assertEquals(2, range.endCol());
    }

    @ParameterizedTest
    @CsvSource({
        "A1:B2, 0, 0, 1, 1",
        "A1, 0, 0, 0, 0",
        "C5:G10, 4, 2, 9, 6",
        "AA1:AB5, 0, 26, 4, 27",
        "Z100:AA200, 99, 25, 199, 26"
    })
    void parseRange(String ref, int startRow, int startCol, int endRow, int endCol) {
        CellRange range = CellRange.parse(ref);

        assertEquals(startRow, range.startRow());
        assertEquals(startCol, range.startCol());
        assertEquals(endRow, range.endRow());
        assertEquals(endCol, range.endCol());
    }

    @Test
    void toRef() {
        CellRange range = CellRange.of(0, 0, 5, 5);
        assertEquals("A1:F6", range.toRef());
    }

    @Test
    void toRefSingleCell() {
        CellRange range = CellRange.of(0, 0, 0, 0);
        assertEquals("A1", range.toRef());
    }

    @Test
    void colCount() {
        CellRange range = CellRange.of(0, 0, 5, 3);
        assertEquals(4, range.colCount());
    }

    @Test
    void rowCount() {
        CellRange range = CellRange.of(2, 0, 10, 0);
        assertEquals(9, range.rowCount());
    }

    @Test
    void contains() {
        CellRange range = CellRange.of(0, 0, 5, 5);

        assertTrue(range.contains(0, 0));
        assertTrue(range.contains(3, 3));
        assertTrue(range.contains(5, 5));
        assertFalse(range.contains(6, 6));
        assertFalse(range.contains(-1, 0));
    }

    @Test
    void invalidRangeThrows() {
        assertThrows(IllegalArgumentException.class, () -> CellRange.of(5, 5, 0, 0));
        assertThrows(IllegalArgumentException.class, () -> CellRange.of(-1, 0, 5, 5));
        assertThrows(IllegalArgumentException.class, () -> CellRange.of(0, -1, 5, 5));
    }

    @Test
    void equality() {
        CellRange range1 = CellRange.of(0, 0, 5, 5);
        CellRange range2 = CellRange.of(0, 0, 5, 5);
        CellRange range3 = CellRange.of(0, 0, 5, 6);

        assertEquals(range1, range2);
        assertEquals(range1.hashCode(), range2.hashCode());
        assertNotEquals(range1, range3);
    }

    @Test
    void toStringReturnsRef() {
        CellRange range = CellRange.of(0, 0, 5, 5);
        assertEquals("A1:F6", range.toString());
    }
}
