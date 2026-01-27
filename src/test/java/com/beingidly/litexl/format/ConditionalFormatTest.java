package com.beingidly.litexl.format;

import com.beingidly.litexl.CellRange;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class ConditionalFormatTest {
    private static final CellRange RANGE = CellRange.parse("A1:B10");
    private static final int STYLE_ID = 1;

    @Test
    void greaterThanCreatesCorrectFormat() {
        var cf = ConditionalFormat.greaterThan(RANGE, 10.0, STYLE_ID);
        assertEquals(ConditionalFormat.Type.CELL_VALUE, cf.type());
        assertEquals(ConditionalFormat.Operator.GREATER_THAN, cf.operator());
        assertEquals("10.0", cf.formula1());
        assertNull(cf.formula2());
        assertEquals(RANGE, cf.range());
        assertEquals(STYLE_ID, cf.styleId());
    }

    @Test
    void lessThanCreatesCorrectFormat() {
        var cf = ConditionalFormat.lessThan(RANGE, 5.0, STYLE_ID);
        assertEquals(ConditionalFormat.Type.CELL_VALUE, cf.type());
        assertEquals(ConditionalFormat.Operator.LESS_THAN, cf.operator());
        assertEquals("5.0", cf.formula1());
        assertNull(cf.formula2());
        assertEquals(RANGE, cf.range());
        assertEquals(STYLE_ID, cf.styleId());
    }

    @Test
    void betweenCreatesCorrectFormat() {
        var cf = ConditionalFormat.between(RANGE, 1.0, 100.0, STYLE_ID);
        assertEquals(ConditionalFormat.Type.CELL_VALUE, cf.type());
        assertEquals(ConditionalFormat.Operator.BETWEEN, cf.operator());
        assertEquals("1.0", cf.formula1());
        assertEquals("100.0", cf.formula2());
        assertEquals(RANGE, cf.range());
        assertEquals(STYLE_ID, cf.styleId());
    }

    @Test
    void expressionCreatesCorrectFormat() {
        var cf = ConditionalFormat.expression(RANGE, "A1>B1", STYLE_ID);
        assertEquals(ConditionalFormat.Type.EXPRESSION, cf.type());
        assertEquals(ConditionalFormat.Operator.NONE, cf.operator());
        assertEquals("A1>B1", cf.formula1());
        assertNull(cf.formula2());
        assertEquals(RANGE, cf.range());
        assertEquals(STYLE_ID, cf.styleId());
    }

    @Test
    void duplicateValuesCreatesCorrectFormat() {
        var cf = ConditionalFormat.duplicateValues(RANGE, STYLE_ID);
        assertEquals(ConditionalFormat.Type.DUPLICATE_VALUES, cf.type());
        assertEquals(ConditionalFormat.Operator.NONE, cf.operator());
        assertNull(cf.formula1());
        assertNull(cf.formula2());
        assertEquals(RANGE, cf.range());
        assertEquals(STYLE_ID, cf.styleId());
    }

    @Test
    void uniqueValuesCreatesCorrectFormat() {
        var cf = ConditionalFormat.uniqueValues(RANGE, STYLE_ID);
        assertEquals(ConditionalFormat.Type.UNIQUE_VALUES, cf.type());
        assertEquals(ConditionalFormat.Operator.NONE, cf.operator());
        assertNull(cf.formula1());
        assertNull(cf.formula2());
        assertEquals(RANGE, cf.range());
        assertEquals(STYLE_ID, cf.styleId());
    }

    @Test
    void allTypesExist() {
        assertEquals(15, ConditionalFormat.Type.values().length);
    }

    @Test
    void allOperatorsExist() {
        assertEquals(9, ConditionalFormat.Operator.values().length);
    }

    @Test
    void recordEquality() {
        var cf1 = ConditionalFormat.greaterThan(RANGE, 10.0, STYLE_ID);
        var cf2 = ConditionalFormat.greaterThan(RANGE, 10.0, STYLE_ID);
        assertEquals(cf1, cf2);
        assertEquals(cf1.hashCode(), cf2.hashCode());
    }

    @Test
    void recordInequalityDifferentValue() {
        var cf1 = ConditionalFormat.greaterThan(RANGE, 10.0, STYLE_ID);
        var cf2 = ConditionalFormat.greaterThan(RANGE, 20.0, STYLE_ID);
        assertNotEquals(cf1, cf2);
    }

    @Test
    void recordInequalityDifferentRange() {
        var range2 = CellRange.parse("C1:D10");
        var cf1 = ConditionalFormat.greaterThan(RANGE, 10.0, STYLE_ID);
        var cf2 = ConditionalFormat.greaterThan(range2, 10.0, STYLE_ID);
        assertNotEquals(cf1, cf2);
    }

    @Test
    void recordInequalityDifferentStyleId() {
        var cf1 = ConditionalFormat.greaterThan(RANGE, 10.0, 1);
        var cf2 = ConditionalFormat.greaterThan(RANGE, 10.0, 2);
        assertNotEquals(cf1, cf2);
    }

    @Test
    void typeEnumValues() {
        // Verify all expected types exist
        assertNotNull(ConditionalFormat.Type.CELL_VALUE);
        assertNotNull(ConditionalFormat.Type.EXPRESSION);
        assertNotNull(ConditionalFormat.Type.COLOR_SCALE);
        assertNotNull(ConditionalFormat.Type.DATA_BAR);
        assertNotNull(ConditionalFormat.Type.ICON_SET);
        assertNotNull(ConditionalFormat.Type.TOP_BOTTOM);
        assertNotNull(ConditionalFormat.Type.ABOVE_AVERAGE);
        assertNotNull(ConditionalFormat.Type.DUPLICATE_VALUES);
        assertNotNull(ConditionalFormat.Type.UNIQUE_VALUES);
        assertNotNull(ConditionalFormat.Type.CONTAINS_TEXT);
        assertNotNull(ConditionalFormat.Type.NOT_CONTAINS_TEXT);
        assertNotNull(ConditionalFormat.Type.BEGINS_WITH);
        assertNotNull(ConditionalFormat.Type.ENDS_WITH);
        assertNotNull(ConditionalFormat.Type.CONTAINS_BLANKS);
        assertNotNull(ConditionalFormat.Type.CONTAINS_ERRORS);
    }

    @Test
    void operatorEnumValues() {
        // Verify all expected operators exist
        assertNotNull(ConditionalFormat.Operator.NONE);
        assertNotNull(ConditionalFormat.Operator.LESS_THAN);
        assertNotNull(ConditionalFormat.Operator.LESS_THAN_OR_EQUAL);
        assertNotNull(ConditionalFormat.Operator.EQUAL);
        assertNotNull(ConditionalFormat.Operator.NOT_EQUAL);
        assertNotNull(ConditionalFormat.Operator.GREATER_THAN_OR_EQUAL);
        assertNotNull(ConditionalFormat.Operator.GREATER_THAN);
        assertNotNull(ConditionalFormat.Operator.BETWEEN);
        assertNotNull(ConditionalFormat.Operator.NOT_BETWEEN);
    }

    @Test
    void greaterThanWithIntegerValue() {
        var cf = ConditionalFormat.greaterThan(RANGE, 100, STYLE_ID);
        assertEquals("100.0", cf.formula1());
    }

    @Test
    void betweenWithNegativeValues() {
        var cf = ConditionalFormat.between(RANGE, -50.0, 50.0, STYLE_ID);
        assertEquals("-50.0", cf.formula1());
        assertEquals("50.0", cf.formula2());
    }

    @Test
    void expressionWithComplexFormula() {
        var cf = ConditionalFormat.expression(RANGE, "AND(A1>0,B1<100)", STYLE_ID);
        assertEquals("AND(A1>0,B1<100)", cf.formula1());
    }

    @Test
    void directRecordConstruction() {
        var cf = new ConditionalFormat(
            RANGE,
            ConditionalFormat.Type.CELL_VALUE,
            ConditionalFormat.Operator.EQUAL,
            "42",
            null,
            STYLE_ID
        );
        assertEquals(ConditionalFormat.Type.CELL_VALUE, cf.type());
        assertEquals(ConditionalFormat.Operator.EQUAL, cf.operator());
        assertEquals("42", cf.formula1());
        assertNull(cf.formula2());
    }
}
