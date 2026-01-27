package com.beingidly.litexl;

import org.junit.jupiter.api.Test;

import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.*;

class CellValueTest {

    @Test
    void emptyValue() {
        CellValue empty = new CellValue.Empty();

        assertEquals(CellType.EMPTY, empty.type());
        assertInstanceOf(CellValue.Empty.class, empty);
    }

    @Test
    void textValue() {
        CellValue text = new CellValue.Text("Hello");

        assertEquals(CellType.STRING, text.type());
        assertInstanceOf(CellValue.Text.class, text);
        assertEquals("Hello", ((CellValue.Text) text).value());
    }

    @Test
    void numberValue() {
        CellValue number = new CellValue.Number(42.5);

        assertEquals(CellType.NUMBER, number.type());
        assertInstanceOf(CellValue.Number.class, number);
        assertEquals(42.5, ((CellValue.Number) number).value());
    }

    @Test
    void booleanValue() {
        CellValue bool = new CellValue.Bool(true);

        assertEquals(CellType.BOOLEAN, bool.type());
        assertInstanceOf(CellValue.Bool.class, bool);
        assertTrue(((CellValue.Bool) bool).value());
    }

    @Test
    void dateValue() {
        LocalDateTime date = LocalDateTime.of(2024, 1, 15, 10, 30);
        CellValue dateVal = new CellValue.Date(date);

        assertEquals(CellType.DATE, dateVal.type());
        assertInstanceOf(CellValue.Date.class, dateVal);
        assertEquals(date, ((CellValue.Date) dateVal).value());
    }

    @Test
    void formulaValue() {
        CellValue formula = new CellValue.Formula("SUM(A1:A10)");

        assertEquals(CellType.FORMULA, formula.type());
        assertInstanceOf(CellValue.Formula.class, formula);
        assertEquals("SUM(A1:A10)", ((CellValue.Formula) formula).expression());
    }

    @Test
    void formulaWithCachedValue() {
        CellValue formula = new CellValue.Formula("SUM(A1:A10)", new CellValue.Number(100.0));

        assertEquals(CellType.FORMULA, formula.type());
        CellValue.Formula f = (CellValue.Formula) formula;
        assertEquals("SUM(A1:A10)", f.expression());
        assertNotNull(f.cached());
        assertEquals(CellType.NUMBER, f.cached().type());
    }

    @Test
    void errorValue() {
        CellValue error = new CellValue.Error("#DIV/0!");

        assertEquals(CellType.ERROR, error.type());
        assertInstanceOf(CellValue.Error.class, error);
        assertEquals("#DIV/0!", ((CellValue.Error) error).code());
    }

    @Test
    void patternMatching() {
        CellValue text = new CellValue.Text("Hello");

        String result = switch (text) {
            case CellValue.Empty _ -> "empty";
            case CellValue.Text t -> "text: " + t.value();
            case CellValue.Number n -> "number: " + n.value();
            case CellValue.Bool b -> "bool: " + b.value();
            case CellValue.Date d -> "date: " + d.value();
            case CellValue.Formula f -> "formula: " + f.expression();
            case CellValue.Error e -> "error: " + e.code();
        };

        assertEquals("text: Hello", result);
    }

}
