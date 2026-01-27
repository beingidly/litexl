package com.beingidly.litexl;

import com.beingidly.litexl.format.ConditionalFormat;
import com.beingidly.litexl.format.DataValidation;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class SheetFormatTest {

    private Workbook wb;
    private Sheet sheet;

    @BeforeEach
    void setUp() {
        wb = Workbook.create();
        sheet = wb.addSheet("Test");
    }

    @Test
    void merge_withCoordinates() {
        sheet.merge(0, 0, 2, 2);
        assertEquals(1, sheet.mergedCells().size());
        var merged = sheet.mergedCells().get(0);
        assertEquals(0, merged.startRow());
        assertEquals(0, merged.startCol());
        assertEquals(2, merged.endRow());
        assertEquals(2, merged.endCol());
    }

    @Test
    void merge_withRange() {
        CellRange range = CellRange.of("B2:D5");
        sheet.format().merge(range);
        assertEquals(1, sheet.mergedCells().size());
    }

    @Test
    void unmerge() {
        sheet.merge(0, 0, 2, 2);
        assertEquals(1, sheet.mergedCells().size());
        sheet.format().unmerge(0, 0, 2, 2);
        assertEquals(0, sheet.mergedCells().size());
    }

    @Test
    void mergedRegion_toRange() {
        sheet.merge(1, 2, 3, 4);
        var merged = sheet.mergedCells().get(0);
        CellRange range = merged.toRange();
        assertEquals("C2:E4", range.toRef());
    }

    @Test
    void autoFilter_withCoordinates() {
        sheet.setAutoFilter(0, 0, 10, 5);
        assertNotNull(sheet.autoFilter());
        assertEquals("A1:F11", sheet.autoFilter().range().toRef());
    }

    @Test
    void autoFilter_withRange() {
        CellRange range = CellRange.of("A1:D10");
        sheet.format().setAutoFilter(range);
        assertNotNull(sheet.autoFilter());
        assertEquals(range, sheet.autoFilter().range());
    }

    @Test
    void clearAutoFilter() {
        sheet.setAutoFilter(0, 0, 10, 5);
        assertNotNull(sheet.autoFilter());
        sheet.clearAutoFilter();
        assertNull(sheet.autoFilter());
    }

    @Test
    void conditionalFormat_addAndGet() {
        var cf = ConditionalFormat.greaterThan(CellRange.of("A1:A10"), 100, 0xFFFF0000);
        sheet.format().addConditionalFormat(cf);
        assertEquals(1, sheet.format().conditionalFormats().size());
    }

    @Test
    void clearConditionalFormats() {
        var cf = ConditionalFormat.greaterThan(CellRange.of("A1:A10"), 100, 0xFFFF0000);
        sheet.format().addConditionalFormat(cf);
        sheet.format().clearConditionalFormats();
        assertTrue(sheet.format().conditionalFormats().isEmpty());
    }

    @Test
    void validation_addAndGet() {
        var dv = DataValidation.wholeNumber(CellRange.of("B1:B10"),
            DataValidation.Operator.BETWEEN, "1", "100");
        sheet.format().addValidation(dv);
        assertEquals(1, sheet.format().validations().size());
    }

    @Test
    void clearValidations() {
        var dv = DataValidation.wholeNumber(CellRange.of("B1:B10"),
            DataValidation.Operator.BETWEEN, "1", "100");
        sheet.format().addValidation(dv);
        sheet.format().clearValidations();
        assertTrue(sheet.format().validations().isEmpty());
    }

    @Test
    void columnWidth_setAndGet() {
        sheet.format().setColumnWidth(0, 15.5);
        ColumnWidth width = sheet.format().getColumnWidth(0);
        assertFalse(width.isAuto());
        assertEquals(15.5, width.getValue());
    }

    @Test
    void columnWidth_unset_returnsAuto() {
        ColumnWidth width = sheet.format().getColumnWidth(5);
        assertTrue(width.isAuto());
    }

    @Test
    void columnWidths_map() {
        sheet.format().setColumnWidth(0, 10.0);
        sheet.format().setColumnWidth(2, 20.0);
        var widths = sheet.format().columnWidths();
        assertEquals(2, widths.size());
        assertEquals(10.0, widths.get(0));
        assertEquals(20.0, widths.get(2));
    }

    @Test
    void hidden_defaultFalse() {
        assertFalse(sheet.format().hidden());
    }

    @Test
    void hidden_setAndGet() {
        sheet.format().setHidden(true);
        assertTrue(sheet.format().hidden());
        sheet.format().setHidden(false);
        assertFalse(sheet.format().hidden());
    }

    @Test
    void sheetFormatMergedRegion_equality() {
        // Test SheetFormat.MergedRegion equality via format().mergedCells()
        sheet.merge(0, 0, 2, 2);
        var regions = sheet.format().mergedCells();
        var r1 = regions.get(0);
        var r2 = new SheetFormat.MergedRegion(0, 0, 2, 2);
        assertEquals(r1, r2);
        assertEquals(r1.hashCode(), r2.hashCode());
    }

    @Test
    void sheetMergedRegion_equality() {
        // Test Sheet.MergedRegion equality via sheet.mergedCells()
        sheet.merge(0, 0, 2, 2);
        var regions = sheet.mergedCells();
        var r1 = regions.get(0);
        var r2 = new Sheet.MergedRegion(0, 0, 2, 2);
        assertEquals(r1, r2);
        assertEquals(r1.hashCode(), r2.hashCode());
    }

    @Test
    void sheetFormatMergedRegion_inequality() {
        var r1 = new SheetFormat.MergedRegion(0, 0, 2, 2);
        var r2 = new SheetFormat.MergedRegion(0, 0, 3, 3);
        var r3 = new SheetFormat.MergedRegion(1, 1, 2, 2);
        assertNotEquals(r1, r2);
        assertNotEquals(r1, r3);
        assertNotEquals(r1.hashCode(), r2.hashCode());
    }

    @Test
    void sheetMergedRegion_inequality() {
        var r1 = new Sheet.MergedRegion(0, 0, 2, 2);
        var r2 = new Sheet.MergedRegion(0, 0, 3, 3);
        var r3 = new Sheet.MergedRegion(1, 1, 2, 2);
        assertNotEquals(r1, r2);
        assertNotEquals(r1, r3);
        assertNotEquals(r1.hashCode(), r2.hashCode());
    }

    @Test
    void sheetFormatMergedRegion_toString() {
        var region = new SheetFormat.MergedRegion(1, 2, 3, 4);
        String str = region.toString();
        assertNotNull(str);
        assertTrue(str.contains("1"));
        assertTrue(str.contains("2"));
        assertTrue(str.contains("3"));
        assertTrue(str.contains("4"));
    }

    @Test
    void sheetMergedRegion_toString() {
        var region = new Sheet.MergedRegion(1, 2, 3, 4);
        String str = region.toString();
        assertNotNull(str);
        assertTrue(str.contains("1"));
        assertTrue(str.contains("2"));
        assertTrue(str.contains("3"));
        assertTrue(str.contains("4"));
    }

    @Test
    void sheetMergedRegion_toRange() {
        var region = new Sheet.MergedRegion(1, 2, 3, 4);
        CellRange range = region.toRange();
        assertEquals("C2:E4", range.toRef());
    }

    @Test
    void sheetFormatMergedRegion_toRange() {
        var region = new SheetFormat.MergedRegion(1, 2, 3, 4);
        CellRange range = region.toRange();
        assertEquals("C2:E4", range.toRef());
    }
}
