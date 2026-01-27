package com.beingidly.litexl;

import org.junit.jupiter.api.Test;

import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.*;

class CellTest {

    @Test
    void createCell() {
        Cell cell = new Cell(5);

        assertEquals(5, cell.column());
        assertEquals(CellType.EMPTY, cell.type());
    }

    @Test
    void setStringValue() {
        Cell cell = new Cell(0);

        cell.set("Hello");

        assertEquals(CellType.STRING, cell.type());
        assertEquals("Hello", cell.string());
    }

    @Test
    void setDoubleValue() {
        Cell cell = new Cell(0);

        cell.set(42.5);

        assertEquals(CellType.NUMBER, cell.type());
        assertEquals(42.5, cell.number());
    }

    @Test
    void setIntValue() {
        Cell cell = new Cell(0);

        cell.set((double) 100);

        assertEquals(CellType.NUMBER, cell.type());
        assertEquals(100.0, cell.number());
    }

    @Test
    void setBooleanValue() {
        Cell cell = new Cell(0);

        cell.set(true);

        assertEquals(CellType.BOOLEAN, cell.type());
        assertTrue(cell.bool());
    }

    @Test
    void setDateTimeValue() {
        Cell cell = new Cell(0);
        LocalDateTime date = LocalDateTime.of(2024, 1, 15, 10, 30);

        cell.set(date);

        assertEquals(CellType.DATE, cell.type());
        assertEquals(date, cell.date());
    }

    @Test
    void setFormula() {
        Cell cell = new Cell(0);

        cell.setFormula("SUM(A1:A10)");

        assertEquals(CellType.FORMULA, cell.type());
        assertEquals("SUM(A1:A10)", cell.formula());
    }

    @Test
    void setValue() {
        Cell cell = new Cell(0);

        cell.setValue(new CellValue.Text("Test"));

        assertEquals(CellType.STRING, cell.type());
        assertEquals("Test", cell.string());
    }

    @Test
    void setEmpty() {
        Cell cell = new Cell(0);
        cell.set("Hello");

        assertEquals(CellType.STRING, cell.type());

        cell.setEmpty();

        assertEquals(CellType.EMPTY, cell.type());
    }

    @Test
    void styleId() {
        Cell cell = new Cell(0);

        assertEquals(0, cell.styleId());

        cell.style(5);

        assertEquals(5, cell.styleId());
    }

    @Test
    void chaining() {
        Cell cell = new Cell(0);

        Cell result = cell.set("Test").style(1);

        assertSame(cell, result);
        assertEquals("Test", cell.string());
        assertEquals(1, cell.styleId());
    }

    @Test
    void getStringOnNonString() {
        Cell cell = new Cell(0);
        cell.set(42.0);

        // Returns null instead of throwing
        assertNull(cell.string());
    }

    @Test
    void getNumberOnNonNumber() {
        Cell cell = new Cell(0);
        cell.set("text");

        // Returns 0 instead of throwing
        assertEquals(0.0, cell.number());
    }

    @Test
    void getBoolOnNonBool() {
        Cell cell = new Cell(0);
        cell.set("text");

        // Returns false instead of throwing
        assertFalse(cell.bool());
    }

    @Test
    void getDateOnNonDate() {
        Cell cell = new Cell(0);
        cell.set("text");

        // Returns null instead of throwing
        assertNull(cell.date());
    }

    @Test
    void getFormulaOnNonFormula() {
        Cell cell = new Cell(0);
        cell.set("text");

        // Returns null instead of throwing
        assertNull(cell.formula());
    }

    @Test
    void valueAccessor() {
        Cell cell = new Cell(0);
        cell.set("Hello");

        CellValue value = cell.value();

        assertInstanceOf(CellValue.Text.class, value);
        assertEquals("Hello", ((CellValue.Text) value).value());
    }

    @Test
    void rawValue() {
        Cell cell = new Cell(0);

        cell.set("Hello");
        assertEquals("Hello", cell.rawValue());

        cell.set(42.5);
        assertEquals(42.5, cell.rawValue());

        cell.set(true);
        assertEquals(true, cell.rawValue());

        cell.setEmpty();
        assertNull(cell.rawValue());
    }

    @Test
    void errorCell() {
        Cell cell = new Cell(0);
        cell.setValue(new CellValue.Error("#DIV/0!"));

        assertEquals(CellType.ERROR, cell.type());
        assertEquals("#DIV/0!", cell.error());
    }

    @Test
    void toStringFormat() {
        Cell cell = new Cell(3);
        cell.set("Hello");

        String str = cell.toString();
        assertTrue(str.contains("3"));
        assertTrue(str.contains("Hello"));
    }

    @Test
    void setNullString() {
        Cell cell = new Cell(0);
        cell.set("test");

        cell.set((String) null);

        assertEquals(CellType.EMPTY, cell.type());
    }

    @Test
    void setNullDateTime() {
        Cell cell = new Cell(0);
        cell.set(LocalDateTime.now());

        cell.set((LocalDateTime) null);

        assertEquals(CellType.EMPTY, cell.type());
    }

    @Test
    void setNullFormula() {
        Cell cell = new Cell(0);
        cell.setFormula("SUM(A1:A10)");

        cell.setFormula(null);

        assertEquals(CellType.EMPTY, cell.type());
    }

    @Test
    void setNullValue() {
        Cell cell = new Cell(0);
        cell.set("test");

        cell.setValue(null);

        assertEquals(CellType.EMPTY, cell.type());
    }
}
