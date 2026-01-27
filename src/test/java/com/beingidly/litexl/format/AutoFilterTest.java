package com.beingidly.litexl.format;

import com.beingidly.litexl.CellRange;
import org.junit.jupiter.api.Test;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class AutoFilterTest {

    @Test
    void autoFilter_hasRangeAndColumns() {
        CellRange range = CellRange.of("A1:D10");
        AutoFilter filter = new AutoFilter(range, List.of());
        assertEquals(range, filter.range());
        assertTrue(filter.columns().isEmpty());
    }

    @Test
    void filterColumn_values() {
        var fc = AutoFilter.FilterColumn.values(2, List.of("Apple", "Banana"));
        assertEquals(2, fc.columnIndex());
        assertEquals(List.of("Apple", "Banana"), fc.values());
        assertNull(fc.custom());
    }

    @Test
    void filterColumn_custom() {
        var custom = AutoFilter.CustomFilter.single(
            AutoFilter.CustomFilter.Operator.EQUAL, "test");
        var fc = AutoFilter.FilterColumn.custom(3, custom);
        assertEquals(3, fc.columnIndex());
        assertTrue(fc.values().isEmpty());
        assertNotNull(fc.custom());
    }

    @Test
    void customFilter_single() {
        var cf = AutoFilter.CustomFilter.single(
            AutoFilter.CustomFilter.Operator.GREATER_THAN, "100");
        assertEquals(AutoFilter.CustomFilter.Operator.GREATER_THAN, cf.op1());
        assertEquals("100", cf.val1());
        assertNull(cf.op2());
        assertNull(cf.val2());
        assertTrue(cf.and());
    }

    @Test
    void customFilter_and() {
        var cf = AutoFilter.CustomFilter.and(
            AutoFilter.CustomFilter.Operator.GREATER_THAN, "10",
            AutoFilter.CustomFilter.Operator.LESS_THAN, "100");
        assertEquals(AutoFilter.CustomFilter.Operator.GREATER_THAN, cf.op1());
        assertEquals("10", cf.val1());
        assertEquals(AutoFilter.CustomFilter.Operator.LESS_THAN, cf.op2());
        assertEquals("100", cf.val2());
        assertTrue(cf.and());
    }

    @Test
    void customFilter_or() {
        var cf = AutoFilter.CustomFilter.or(
            AutoFilter.CustomFilter.Operator.EQUAL, "A",
            AutoFilter.CustomFilter.Operator.EQUAL, "B");
        assertFalse(cf.and());
    }

    @Test
    void customFilter_operatorValues() {
        assertEquals(6, AutoFilter.CustomFilter.Operator.values().length);
        assertNotNull(AutoFilter.CustomFilter.Operator.valueOf("EQUAL"));
        assertNotNull(AutoFilter.CustomFilter.Operator.valueOf("NOT_EQUAL"));
        assertNotNull(AutoFilter.CustomFilter.Operator.valueOf("GREATER_THAN"));
        assertNotNull(AutoFilter.CustomFilter.Operator.valueOf("GREATER_THAN_OR_EQUAL"));
        assertNotNull(AutoFilter.CustomFilter.Operator.valueOf("LESS_THAN"));
        assertNotNull(AutoFilter.CustomFilter.Operator.valueOf("LESS_THAN_OR_EQUAL"));
    }
}
