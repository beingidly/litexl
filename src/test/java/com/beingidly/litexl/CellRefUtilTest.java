package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.*;

class CellRefUtilTest {

    @ParameterizedTest
    @CsvSource({
        "0, A",
        "1, B",
        "25, Z",
        "26, AA",
        "27, AB",
        "51, AZ",
        "52, BA",
        "701, ZZ",
        "702, AAA"
    })
    void colToLetters(int col, String expected) {
        assertEquals(expected, CellRefUtil.colToLetters(col));
    }

    @ParameterizedTest
    @CsvSource({
        "A, 0",
        "B, 1",
        "Z, 25",
        "AA, 26",
        "AB, 27",
        "AZ, 51",
        "BA, 52",
        "ZZ, 701",
        "AAA, 702"
    })
    void lettersToCol(String letters, int expected) {
        assertEquals(expected, CellRefUtil.lettersToCol(letters));
    }

    @Test
    void toRef() {
        assertEquals("A1", CellRefUtil.toRef(0, 0));
        assertEquals("B2", CellRefUtil.toRef(1, 1));
        assertEquals("Z100", CellRefUtil.toRef(99, 25));
        assertEquals("AA1", CellRefUtil.toRef(0, 26));
    }

    @Test
    void parseRef() {
        assertArrayEquals(new int[]{0, 0}, CellRefUtil.parseRef("A1"));
        assertArrayEquals(new int[]{1, 1}, CellRefUtil.parseRef("B2"));
        assertArrayEquals(new int[]{99, 25}, CellRefUtil.parseRef("Z100"));
        assertArrayEquals(new int[]{0, 26}, CellRefUtil.parseRef("AA1"));
    }

    @Test
    void parseRefCaseInsensitive() {
        assertArrayEquals(new int[]{0, 0}, CellRefUtil.parseRef("a1"));
        assertArrayEquals(new int[]{0, 26}, CellRefUtil.parseRef("aa1"));
    }

    @Test
    void invalidRef() {
        assertThrows(IllegalArgumentException.class, () -> CellRefUtil.parseRef(""));
        assertThrows(IllegalArgumentException.class, () -> CellRefUtil.parseRef("1"));
        assertThrows(IllegalArgumentException.class, () -> CellRefUtil.parseRef("A"));
        assertThrows(IllegalArgumentException.class, () -> CellRefUtil.parseRef("A0"));
    }

    @Test
    void invalidCol() {
        assertThrows(IllegalArgumentException.class, () -> CellRefUtil.colToLetters(-1));
        assertThrows(IllegalArgumentException.class, () -> CellRefUtil.lettersToCol(""));
    }
}
