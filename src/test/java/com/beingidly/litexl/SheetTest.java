package com.beingidly.litexl;

import com.beingidly.litexl.crypto.SheetProtection;
import com.beingidly.litexl.format.ConditionalFormat;
import com.beingidly.litexl.format.DataValidation;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class SheetTest {

    @Test
    void createSheet() {
        Sheet sheet = new Sheet("Test", 0);

        assertEquals("Test", sheet.name());
        assertEquals(0, sheet.index());
        assertEquals(0, sheet.rowCount());
    }

    @Test
    void setName() {
        Sheet sheet = new Sheet("Test", 0);
        sheet.setName("NewName");

        assertEquals("NewName", sheet.name());
    }

    @Test
    void setNameEmptyThrows() {
        Sheet sheet = new Sheet("Test", 0);

        assertThrows(IllegalArgumentException.class, () -> sheet.setName(""));
    }

    @Test
    void cellCreation() {
        Sheet sheet = new Sheet("Test", 0);

        Cell cell = sheet.cell(0, 0);
        assertNotNull(cell);
        assertEquals(0, cell.column());

        // Verify row was created
        assertEquals(1, sheet.rowCount());
    }

    @Test
    void getCell() {
        Sheet sheet = new Sheet("Test", 0);

        // Cell doesn't exist yet
        assertNull(sheet.getCell(0, 0));

        // Create cell
        sheet.cell(0, 0).set("test");

        // Now it exists
        assertNotNull(sheet.getCell(0, 0));
        assertEquals("test", sheet.getCell(0, 0).string());
    }

    @Test
    void rowCreation() {
        Sheet sheet = new Sheet("Test", 0);

        Row row = sheet.row(5);
        assertNotNull(row);
        assertEquals(5, row.rowNum());
    }

    @Test
    void getRow() {
        Sheet sheet = new Sheet("Test", 0);

        assertNull(sheet.getRow(5));

        sheet.row(5);

        assertNotNull(sheet.getRow(5));
    }

    @Test
    void firstAndLastRow() {
        Sheet sheet = new Sheet("Test", 0);

        assertFalse(sheet.firstRow().exists());
        assertFalse(sheet.lastRow().exists());

        sheet.row(3);
        sheet.row(10);
        assertThrows(IllegalStateException.class, () -> sheet.row(5));

        assertTrue(sheet.firstRow().exists());
        assertTrue(sheet.lastRow().exists());
        assertEquals(3, sheet.firstRow().getValue());
        assertEquals(10, sheet.lastRow().getValue());
    }

    @Test
    void mergeCells() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.merge(0, 0, 2, 2);

        assertEquals(1, sheet.mergedCells().size());
        Sheet.MergedRegion merge = sheet.mergedCells().get(0);
        assertEquals(0, merge.startRow());
        assertEquals(0, merge.startCol());
        assertEquals(2, merge.endRow());
        assertEquals(2, merge.endCol());
    }

    @Test
    void mergeCellsWithRange() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.merge(CellRange.of(1, 1, 3, 3));

        assertEquals(1, sheet.mergedCells().size());
    }

    @Test
    void unmergeCells() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.merge(0, 0, 2, 2);
        assertEquals(1, sheet.mergedCells().size());

        sheet.unmerge(0, 0, 2, 2);
        assertEquals(0, sheet.mergedCells().size());
    }

    @Test
    void autoFilter() {
        Sheet sheet = new Sheet("Test", 0);

        assertNull(sheet.autoFilter());

        sheet.setAutoFilter(0, 0, 10, 5);

        assertNotNull(sheet.autoFilter());
        assertEquals(0, sheet.autoFilter().range().startRow());
        assertEquals(5, sheet.autoFilter().range().endCol());
    }

    @Test
    void autoFilterWithRange() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.setAutoFilter(CellRange.of(1, 1, 20, 10));

        assertNotNull(sheet.autoFilter());
        assertEquals(1, sheet.autoFilter().range().startRow());
        assertEquals(10, sheet.autoFilter().range().endCol());
    }

    @Test
    void clearAutoFilter() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.setAutoFilter(0, 0, 10, 5);
        assertNotNull(sheet.autoFilter());

        sheet.clearAutoFilter();
        assertNull(sheet.autoFilter());
    }

    @Test
    void sheetProtection() {
        Sheet sheet = new Sheet("Test", 0);

        assertFalse(sheet.isProtected());
        assertNull(sheet.protection());

        sheet.protectionManager().protect("password".toCharArray(), SheetProtection.defaults());

        assertTrue(sheet.isProtected());
        assertNotNull(sheet.protection());
    }

    @Test
    void sheetProtectionWithoutPassword() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.protect(SheetProtection.defaults());

        assertTrue(sheet.isProtected());
    }

    @Test
    void unprotectSheet() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.protectionManager().protect("password".toCharArray(), SheetProtection.defaults());
        assertTrue(sheet.isProtected());

        sheet.protectionManager().unprotect("password".toCharArray());
        assertFalse(sheet.isProtected());
    }

    @Test
    void protectionManagerWithCharArray() {
        Sheet sheet = new Sheet("Test", 0);

        assertFalse(sheet.isProtected());

        // Use char[] API via protectionManager
        char[] password = "securePassword".toCharArray();
        sheet.protectionManager().protect(password, SheetProtection.defaults());

        assertTrue(sheet.isProtected());
        assertNotNull(sheet.protection());
        // Password array should be cleared after use
        assertEquals('\0', password[0]);
    }

    @Test
    void protectionManagerUnprotect() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.protectionManager().protect("test123".toCharArray(), SheetProtection.defaults());
        assertTrue(sheet.isProtected());

        // Unprotect with correct password
        boolean result = sheet.protectionManager().unprotect("test123".toCharArray());
        assertTrue(result);
        assertFalse(sheet.isProtected());
    }

    @Test
    void protectionManagerUnprotectWrongPassword() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.protectionManager().protect("correct".toCharArray(), SheetProtection.defaults());
        assertTrue(sheet.isProtected());

        // Try to unprotect with wrong password
        boolean result = sheet.protectionManager().unprotect("wrong".toCharArray());
        assertFalse(result);
        assertTrue(sheet.isProtected());
    }

    @Test
    void formatDelegation() {
        Sheet sheet = new Sheet("Test", 0);

        // Test that format() returns the same SheetFormat instance
        SheetFormat format = sheet.format();
        assertNotNull(format);

        // Test that format operations work via delegation
        format.merge(0, 0, 2, 2);
        assertEquals(1, sheet.mergedCells().size());

        format.setColumnWidth(0, 15.5);
        assertFalse(sheet.getColumnWidth(0).isAuto());
        assertEquals(15.5, sheet.getColumnWidth(0).getValue());
    }

    @Test
    void conditionalFormatting() {
        Sheet sheet = new Sheet("Test", 0);

        assertEquals(0, sheet.conditionalFormats().size());

        ConditionalFormat cf = new ConditionalFormat(
            CellRange.of(0, 0, 10, 10),
            ConditionalFormat.Type.CELL_VALUE,
            ConditionalFormat.Operator.GREATER_THAN,
            "100", null, 0
        );

        sheet.addConditionalFormat(cf);

        assertEquals(1, sheet.conditionalFormats().size());
    }

    @Test
    void clearConditionalFormats() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.addConditionalFormat(new ConditionalFormat(
            CellRange.of(0, 0, 10, 10),
            ConditionalFormat.Type.CELL_VALUE,
            ConditionalFormat.Operator.EQUAL,
            "0", null, 0
        ));

        assertEquals(1, sheet.conditionalFormats().size());

        sheet.clearConditionalFormats();

        assertEquals(0, sheet.conditionalFormats().size());
    }

    @Test
    void dataValidation() {
        Sheet sheet = new Sheet("Test", 0);

        assertEquals(0, sheet.validations().size());

        DataValidation dv = DataValidation.list(
            CellRange.of(0, 0, 10, 0),
            "A", "B", "C"
        );

        sheet.addValidation(dv);

        assertEquals(1, sheet.validations().size());
    }

    @Test
    void clearValidations() {
        Sheet sheet = new Sheet("Test", 0);

        sheet.addValidation(DataValidation.wholeNumber(
            CellRange.of(0, 0, 10, 0),
            1, 10
        ));

        assertEquals(1, sheet.validations().size());

        sheet.clearValidations();

        assertEquals(0, sheet.validations().size());
    }

    @Test
    void columnWidth() {
        Sheet sheet = new Sheet("Test", 0);

        assertTrue(sheet.getColumnWidth(0).isAuto());

        sheet.setColumnWidth(0, 15.5);

        assertFalse(sheet.getColumnWidth(0).isAuto());
        assertEquals(15.5, sheet.getColumnWidth(0).getValue());
    }

    @Test
    void hidden() {
        Sheet sheet = new Sheet("Test", 0);

        assertFalse(sheet.hidden());

        sheet.setHidden(true);

        assertTrue(sheet.hidden());
    }

    @Test
    void toStringFormat() {
        Sheet sheet = new Sheet("Test", 0);
        sheet.row(0);
        sheet.row(1);

        String str = sheet.toString();
        assertTrue(str.contains("Test"));
        assertTrue(str.contains("2"));
    }

    @Test
    void mergedRegionToRange() {
        Sheet.MergedRegion merge = new Sheet.MergedRegion(0, 0, 5, 5);
        CellRange range = merge.toRange();

        assertEquals(0, range.startRow());
        assertEquals(0, range.startCol());
        assertEquals(5, range.endRow());
        assertEquals(5, range.endCol());
    }
}
