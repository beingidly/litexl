package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class ExcelLimitsTest {

    @Test
    void constantsHaveCorrectValues() {
        assertEquals(1_048_576, ExcelLimits.MAX_ROWS);
        assertEquals(16_384, ExcelLimits.MAX_COLUMNS);
        assertEquals(1_048_575, ExcelLimits.MAX_ROW_INDEX);
        assertEquals(16_383, ExcelLimits.MAX_COLUMN_INDEX);
    }

    @Test
    void validateRowIndex_validRange() {
        assertDoesNotThrow(() -> ExcelLimits.validateRowIndex(0));
        assertDoesNotThrow(() -> ExcelLimits.validateRowIndex(100));
        assertDoesNotThrow(() -> ExcelLimits.validateRowIndex(ExcelLimits.MAX_ROW_INDEX));
    }

    @Test
    void validateRowIndex_negativeThrows() {
        var ex = assertThrows(IllegalArgumentException.class,
            () -> ExcelLimits.validateRowIndex(-1));
        assertTrue(ex.getMessage().contains("-1"));
    }

    @Test
    void validateRowIndex_tooLargeThrows() {
        var ex = assertThrows(IllegalArgumentException.class,
            () -> ExcelLimits.validateRowIndex(ExcelLimits.MAX_ROWS));
        assertTrue(ex.getMessage().contains(String.valueOf(ExcelLimits.MAX_ROWS)));
    }

    @Test
    void validateColumnIndex_validRange() {
        assertDoesNotThrow(() -> ExcelLimits.validateColumnIndex(0));
        assertDoesNotThrow(() -> ExcelLimits.validateColumnIndex(100));
        assertDoesNotThrow(() -> ExcelLimits.validateColumnIndex(ExcelLimits.MAX_COLUMN_INDEX));
    }

    @Test
    void validateColumnIndex_negativeThrows() {
        var ex = assertThrows(IllegalArgumentException.class,
            () -> ExcelLimits.validateColumnIndex(-1));
        assertTrue(ex.getMessage().contains("-1"));
    }

    @Test
    void validateColumnIndex_tooLargeThrows() {
        var ex = assertThrows(IllegalArgumentException.class,
            () -> ExcelLimits.validateColumnIndex(ExcelLimits.MAX_COLUMNS));
        assertTrue(ex.getMessage().contains(String.valueOf(ExcelLimits.MAX_COLUMNS)));
    }

    @Test
    void validateCellIndex_valid() {
        assertDoesNotThrow(() -> ExcelLimits.validateCellIndex(0, 0));
        assertDoesNotThrow(() -> ExcelLimits.validateCellIndex(
            ExcelLimits.MAX_ROW_INDEX, ExcelLimits.MAX_COLUMN_INDEX));
    }

    @Test
    void validateCellIndex_invalidRow() {
        assertThrows(IllegalArgumentException.class,
            () -> ExcelLimits.validateCellIndex(-1, 0));
    }

    @Test
    void validateCellIndex_invalidColumn() {
        assertThrows(IllegalArgumentException.class,
            () -> ExcelLimits.validateCellIndex(0, -1));
    }
}
