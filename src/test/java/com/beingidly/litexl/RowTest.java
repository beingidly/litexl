package com.beingidly.litexl;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class RowTest {

    @Test
    void createRow() {
        Row row = new Row(5);

        assertEquals(5, row.rowNum());
        assertTrue(row.cells().isEmpty());
    }

    @Test
    void cellCreation() {
        Row row = new Row(0);

        Cell cell = row.cell(3);

        assertNotNull(cell);
        assertEquals(3, cell.column());
        assertEquals(1, row.cells().size());
    }

    @Test
    void getCell() {
        Row row = new Row(0);

        assertNull(row.getCell(0));

        row.cell(0).set("test");

        assertNotNull(row.getCell(0));
    }

    @Test
    void firstAndLastColumn() {
        Row row = new Row(0);

        assertFalse(row.firstColumn().exists());
        assertFalse(row.lastColumn().exists());

        row.cell(5);
        row.cell(2);
        row.cell(10);

        assertTrue(row.firstColumn().exists());
        assertTrue(row.lastColumn().exists());
        assertEquals(2, row.firstColumn().getValue());
        assertEquals(10, row.lastColumn().getValue());
    }

    @Test
    void customHeight() {
        Row row = new Row(0);

        assertFalse(row.hasCustomHeight());

        row.height(25.0);

        assertTrue(row.hasCustomHeight());
        assertEquals(25.0, row.height());
    }

    @Test
    void defaultHeight() {
        Row row = new Row(0);

        assertEquals(-1, row.height()); // Default height (auto)
    }

    @Test
    void hidden() {
        Row row = new Row(0);

        assertFalse(row.hidden());

        row.hidden(true);

        assertTrue(row.hidden());
    }

    @Test
    void toStringFormat() {
        Row row = new Row(5);
        row.cell(0);
        row.cell(1);

        String str = row.toString();
        assertTrue(str.contains("5"));
        assertTrue(str.contains("2"));
    }

    @Test
    void cellCount() {
        Row row = new Row(0);

        assertEquals(0, row.cellCount());

        row.cell(0);
        row.cell(5);
        row.cell(10);

        assertEquals(3, row.cellCount());
    }

    @Test
    void heightChaining() {
        Row row = new Row(0);

        Row result = row.height(20.0);

        assertSame(row, result);
        assertEquals(20.0, row.height());
    }

    @Test
    void hiddenChaining() {
        Row row = new Row(0);

        Row result = row.hidden(true);

        assertSame(row, result);
        assertTrue(row.hidden());
    }
}
