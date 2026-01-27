package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.CellValue;
import org.junit.jupiter.api.Test;
import java.time.LocalDate;
import java.time.LocalDateTime;
import static org.junit.jupiter.api.Assertions.*;

class TypeConverterTest {

    // ===== String conversions =====

    @Test
    void stringFromCell() {
        assertEquals("hello", TypeConverter.fromCell(new CellValue.Text("hello"), String.class));
        assertNull(TypeConverter.fromCell(new CellValue.Empty(), String.class));
    }

    @Test
    void stringToCell() {
        assertEquals(new CellValue.Text("hello"), TypeConverter.toCell("hello", String.class));
        assertEquals(new CellValue.Empty(), TypeConverter.toCell(null, String.class));
    }

    @Test
    void stringFromNumberCell() {
        // Number to String conversion
        assertEquals("42.0", TypeConverter.fromCell(new CellValue.Number(42.0), String.class));
        assertEquals("3.14", TypeConverter.fromCell(new CellValue.Number(3.14), String.class));
    }

    @Test
    void stringFromBoolCell() {
        // Boolean to String conversion
        assertEquals("true", TypeConverter.fromCell(new CellValue.Bool(true), String.class));
        assertEquals("false", TypeConverter.fromCell(new CellValue.Bool(false), String.class));
    }

    @Test
    void stringFromDateCell() {
        // Date to String conversion without format
        var dt = LocalDateTime.of(2024, 1, 15, 10, 30);
        assertEquals("2024-01-15T10:30", TypeConverter.fromCell(new CellValue.Date(dt), String.class));
    }

    @Test
    void stringFromDateCellWithFormat() {
        // Date to String conversion with custom format
        var dt = LocalDateTime.of(2024, 1, 15, 10, 30);
        assertEquals("2024-01-15", TypeConverter.fromCell(new CellValue.Date(dt), String.class, "yyyy-MM-dd"));
        assertEquals("15/01/2024", TypeConverter.fromCell(new CellValue.Date(dt), String.class, "dd/MM/yyyy"));
    }

    // ===== Integer conversions =====

    @Test
    void intFromCell() {
        assertEquals(42, TypeConverter.fromCell(new CellValue.Number(42.0), int.class));
        assertEquals(42, TypeConverter.fromCell(new CellValue.Number(42.0), Integer.class));
    }

    @Test
    void intToCell() {
        assertEquals(new CellValue.Number(42), TypeConverter.toCell(42, int.class));
        assertEquals(new CellValue.Number(42), TypeConverter.toCell(42, Integer.class));
    }

    @Test
    void intFromNonNumberCell() {
        // Text cell to int should return null
        assertNull(TypeConverter.fromCell(new CellValue.Text("42"), Integer.class));
        assertNull(TypeConverter.fromCell(new CellValue.Bool(true), Integer.class));
    }

    // ===== Long conversions =====

    @Test
    void longFromCell() {
        assertEquals(100L, TypeConverter.fromCell(new CellValue.Number(100.0), long.class));
        assertEquals(100L, TypeConverter.fromCell(new CellValue.Number(100.0), Long.class));
    }

    @Test
    void longToCell() {
        assertEquals(new CellValue.Number(100L), TypeConverter.toCell(100L, long.class));
        assertEquals(new CellValue.Number(100L), TypeConverter.toCell(100L, Long.class));
    }

    @Test
    void longFromNonNumberCell() {
        assertNull(TypeConverter.fromCell(new CellValue.Text("100"), Long.class));
    }

    // ===== Double conversions =====

    @Test
    void doubleFromCell() {
        assertEquals(3.14, TypeConverter.fromCell(new CellValue.Number(3.14), double.class));
        assertEquals(3.14, TypeConverter.fromCell(new CellValue.Number(3.14), Double.class));
    }

    @Test
    void doubleToCell() {
        assertEquals(new CellValue.Number(3.14), TypeConverter.toCell(3.14, double.class));
        assertEquals(new CellValue.Number(3.14), TypeConverter.toCell(3.14, Double.class));
    }

    @Test
    void doubleFromNonNumberCell() {
        assertNull(TypeConverter.fromCell(new CellValue.Text("3.14"), Double.class));
    }

    // ===== Float conversions =====

    @Test
    void floatFromCell() {
        assertEquals(2.5f, TypeConverter.fromCell(new CellValue.Number(2.5), float.class));
        assertEquals(2.5f, TypeConverter.fromCell(new CellValue.Number(2.5), Float.class));
    }

    @Test
    void floatToCell() {
        assertEquals(new CellValue.Number(2.5f), TypeConverter.toCell(2.5f, float.class));
        assertEquals(new CellValue.Number(2.5f), TypeConverter.toCell(2.5f, Float.class));
    }

    @Test
    void floatFromNonNumberCell() {
        assertNull(TypeConverter.fromCell(new CellValue.Text("2.5"), Float.class));
    }

    // ===== Boolean conversions =====

    @Test
    void booleanFromCell() {
        assertEquals(true, TypeConverter.fromCell(new CellValue.Bool(true), boolean.class));
        assertEquals(false, TypeConverter.fromCell(new CellValue.Bool(false), Boolean.class));
    }

    @Test
    void booleanToCell() {
        assertEquals(new CellValue.Bool(true), TypeConverter.toCell(true, boolean.class));
        assertEquals(new CellValue.Bool(false), TypeConverter.toCell(false, Boolean.class));
    }

    @Test
    void booleanFromNonBoolCell() {
        assertNull(TypeConverter.fromCell(new CellValue.Text("true"), Boolean.class));
        assertNull(TypeConverter.fromCell(new CellValue.Number(1.0), Boolean.class));
    }

    // ===== LocalDateTime conversions =====

    @Test
    void localDateTimeFromCell() {
        var dt = LocalDateTime.of(2024, 1, 15, 10, 30);
        assertEquals(dt, TypeConverter.fromCell(new CellValue.Date(dt), LocalDateTime.class));
    }

    @Test
    void localDateTimeToCell() {
        var dt = LocalDateTime.of(2024, 1, 15, 10, 30);
        var expected = new CellValue.Date(dt);
        assertEquals(expected, TypeConverter.toCell(dt, LocalDateTime.class));
    }

    @Test
    void localDateTimeFromNonDateCell() {
        assertNull(TypeConverter.fromCell(new CellValue.Text("2024-01-15"), LocalDateTime.class));
        assertNull(TypeConverter.fromCell(new CellValue.Number(45306.0), LocalDateTime.class));
    }

    // ===== LocalDate conversions =====

    @Test
    void localDateFromCell() {
        var dt = LocalDateTime.of(2024, 1, 15, 0, 0);
        assertEquals(LocalDate.of(2024, 1, 15), TypeConverter.fromCell(new CellValue.Date(dt), LocalDate.class));
    }

    @Test
    void localDateToCell() {
        var date = LocalDate.of(2024, 1, 15);
        var expected = new CellValue.Date(LocalDateTime.of(2024, 1, 15, 0, 0));
        assertEquals(expected, TypeConverter.toCell(date, LocalDate.class));
    }

    @Test
    void localDateFromNonDateCell() {
        assertNull(TypeConverter.fromCell(new CellValue.Text("2024-01-15"), LocalDate.class));
    }

    // ===== Empty cell handling for primitive types (default values) =====

    @Test
    void emptyCellReturnsDefaultForPrimitiveInt() {
        assertEquals(0, TypeConverter.fromCell(new CellValue.Empty(), int.class));
    }

    @Test
    void emptyCellReturnsDefaultForPrimitiveLong() {
        assertEquals(0L, TypeConverter.fromCell(new CellValue.Empty(), long.class));
    }

    @Test
    void emptyCellReturnsDefaultForPrimitiveDouble() {
        assertEquals(0.0, TypeConverter.fromCell(new CellValue.Empty(), double.class));
    }

    @Test
    void emptyCellReturnsDefaultForPrimitiveFloat() {
        assertEquals(0.0f, TypeConverter.fromCell(new CellValue.Empty(), float.class));
    }

    @Test
    void emptyCellReturnsDefaultForPrimitiveBoolean() {
        assertEquals(false, TypeConverter.fromCell(new CellValue.Empty(), boolean.class));
    }

    // ===== Empty cell handling for wrapper types (null) =====

    @Test
    void emptyCellReturnsNullForWrapperInteger() {
        assertNull(TypeConverter.fromCell(new CellValue.Empty(), Integer.class));
    }

    @Test
    void emptyCellReturnsNullForWrapperLong() {
        assertNull(TypeConverter.fromCell(new CellValue.Empty(), Long.class));
    }

    @Test
    void emptyCellReturnsNullForWrapperDouble() {
        assertNull(TypeConverter.fromCell(new CellValue.Empty(), Double.class));
    }

    @Test
    void emptyCellReturnsNullForWrapperFloat() {
        assertNull(TypeConverter.fromCell(new CellValue.Empty(), Float.class));
    }

    @Test
    void emptyCellReturnsNullForWrapperBoolean() {
        assertNull(TypeConverter.fromCell(new CellValue.Empty(), Boolean.class));
    }

    @Test
    void emptyCellReturnsNullForLocalDateTime() {
        assertNull(TypeConverter.fromCell(new CellValue.Empty(), LocalDateTime.class));
    }

    @Test
    void emptyCellReturnsNullForLocalDate() {
        assertNull(TypeConverter.fromCell(new CellValue.Empty(), LocalDate.class));
    }

    // ===== Null value toCell handling =====

    @Test
    void nullValueToCellReturnsEmpty() {
        assertEquals(new CellValue.Empty(), TypeConverter.toCell(null, Integer.class));
        assertEquals(new CellValue.Empty(), TypeConverter.toCell(null, Long.class));
        assertEquals(new CellValue.Empty(), TypeConverter.toCell(null, Double.class));
        assertEquals(new CellValue.Empty(), TypeConverter.toCell(null, Float.class));
        assertEquals(new CellValue.Empty(), TypeConverter.toCell(null, Boolean.class));
        assertEquals(new CellValue.Empty(), TypeConverter.toCell(null, LocalDateTime.class));
        assertEquals(new CellValue.Empty(), TypeConverter.toCell(null, LocalDate.class));
    }

    // ===== Unsupported type handling =====

    @Test
    void fromCellThrowsForUnsupportedType() {
        var exception = assertThrows(IllegalArgumentException.class, () ->
            TypeConverter.fromCell(new CellValue.Text("test"), Object.class)
        );
        assertTrue(exception.getMessage().contains("Unsupported type"));
    }

    @Test
    void toCellThrowsForUnsupportedType() {
        var exception = assertThrows(IllegalArgumentException.class, () ->
            TypeConverter.toCell("test", Object.class)
        );
        assertTrue(exception.getMessage().contains("Unsupported type"));
    }

    // ===== isSupported =====

    @Test
    void isSupported() {
        assertTrue(TypeConverter.isSupported(String.class));
        assertTrue(TypeConverter.isSupported(int.class));
        assertTrue(TypeConverter.isSupported(Integer.class));
        assertTrue(TypeConverter.isSupported(long.class));
        assertTrue(TypeConverter.isSupported(Long.class));
        assertTrue(TypeConverter.isSupported(double.class));
        assertTrue(TypeConverter.isSupported(Double.class));
        assertTrue(TypeConverter.isSupported(float.class));
        assertTrue(TypeConverter.isSupported(Float.class));
        assertTrue(TypeConverter.isSupported(boolean.class));
        assertTrue(TypeConverter.isSupported(Boolean.class));
        assertTrue(TypeConverter.isSupported(LocalDateTime.class));
        assertTrue(TypeConverter.isSupported(LocalDate.class));
        assertFalse(TypeConverter.isSupported(Object.class));
    }
}
