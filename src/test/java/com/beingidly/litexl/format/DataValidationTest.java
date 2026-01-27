package com.beingidly.litexl.format;

import com.beingidly.litexl.CellRange;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class DataValidationTest {
    private static final CellRange RANGE = CellRange.of("A1:A10");

    @Test
    void listWithItemsCreatesCorrectValidation() {
        var dv = DataValidation.list(RANGE, "Apple", "Banana", "Cherry");
        assertEquals(DataValidation.Type.LIST, dv.type());
        assertNull(dv.operator());
        assertEquals("\"Apple,Banana,Cherry\"", dv.formula1());
        assertEquals(RANGE, dv.range());
    }

    @Test
    void listWithRangeCreatesCorrectValidation() {
        var sourceRange = CellRange.of("Z1:Z10");
        var dv = DataValidation.list(RANGE, sourceRange);
        assertEquals(DataValidation.Type.LIST, dv.type());
        assertEquals("$Z$1:$Z$10", dv.formula1());
    }

    @Test
    void listEscapesQuotesInItems() {
        var dv = DataValidation.list(RANGE, "Item\"With\"Quotes");
        assertTrue(dv.formula1().contains("\"\""));
    }

    @Test
    void wholeNumberCreatesCorrectValidation() {
        var dv = DataValidation.wholeNumber(RANGE, DataValidation.Operator.BETWEEN, "1", "100");
        assertEquals(DataValidation.Type.WHOLE, dv.type());
        assertEquals(DataValidation.Operator.BETWEEN, dv.operator());
        assertEquals("1", dv.formula1());
        assertEquals("100", dv.formula2());
    }

    @Test
    void decimalCreatesCorrectValidation() {
        var dv = DataValidation.decimal(RANGE, DataValidation.Operator.GREATER_THAN, "0.5", null);
        assertEquals(DataValidation.Type.DECIMAL, dv.type());
        assertEquals(DataValidation.Operator.GREATER_THAN, dv.operator());
    }

    @Test
    void textLengthCreatesCorrectValidation() {
        var dv = DataValidation.textLength(RANGE, DataValidation.Operator.LESS_THAN_OR_EQUAL, "255", null);
        assertEquals(DataValidation.Type.TEXT_LENGTH, dv.type());
    }

    @Test
    void customCreatesCorrectValidation() {
        var dv = DataValidation.custom(RANGE, "AND(A1>0,A1<100)");
        assertEquals(DataValidation.Type.CUSTOM, dv.type());
        assertEquals("AND(A1>0,A1<100)", dv.formula1());
    }

    @Test
    void allTypesExist() {
        assertEquals(8, DataValidation.Type.values().length);
    }

    @Test
    void allOperatorsExist() {
        assertEquals(8, DataValidation.Operator.values().length);
    }

    @Test
    void hasDefaultErrorMessages() {
        var dv = DataValidation.wholeNumber(RANGE, DataValidation.Operator.BETWEEN, "1", "100");
        assertNotNull(dv.errorTitle());
        assertNotNull(dv.errorMessage());
    }
}
