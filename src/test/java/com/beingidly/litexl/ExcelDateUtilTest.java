package com.beingidly.litexl;

import org.junit.jupiter.api.Test;

import java.time.LocalDate;
import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.*;

class ExcelDateUtilTest {

    @Test
    void toExcelDateKnownDates() {
        // January 1, 1900 = 1
        assertEquals(1, ExcelDateUtil.toExcelDate(LocalDate.of(1900, 1, 1)));

        // January 1, 2000 = 36526
        assertEquals(36526, ExcelDateUtil.toExcelDate(LocalDate.of(2000, 1, 1)));

        // January 1, 2024 = 45292
        assertEquals(45292, ExcelDateUtil.toExcelDate(LocalDate.of(2024, 1, 1)));
    }

    @Test
    void fromExcelDateKnownDates() {
        assertEquals(LocalDate.of(1900, 1, 1), ExcelDateUtil.fromExcelDateToDate(1));
        assertEquals(LocalDate.of(2000, 1, 1), ExcelDateUtil.fromExcelDateToDate(36526));
        assertEquals(LocalDate.of(2024, 1, 1), ExcelDateUtil.fromExcelDateToDate(45292));
    }

    @Test
    void dateTimeWithTime() {
        LocalDateTime dt = LocalDateTime.of(2024, 1, 15, 12, 30, 0);
        double excelDate = ExcelDateUtil.toExcelDate(dt);

        LocalDateTime result = ExcelDateUtil.fromExcelDate(excelDate);

        assertEquals(dt.toLocalDate(), result.toLocalDate());
        assertEquals(dt.getHour(), result.getHour());
        assertEquals(dt.getMinute(), result.getMinute());
    }

    @Test
    void roundTrip() {
        LocalDateTime original = LocalDateTime.of(2024, 6, 15, 14, 30, 45);
        double excelDate = ExcelDateUtil.toExcelDate(original);
        LocalDateTime result = ExcelDateUtil.fromExcelDate(excelDate);

        assertEquals(original.toLocalDate(), result.toLocalDate());
        assertEquals(original.getHour(), result.getHour());
        assertEquals(original.getMinute(), result.getMinute());
        // Seconds might have slight rounding
        assertEquals(original.getSecond(), result.getSecond(), 1);
    }

    @Test
    void isValidExcelDate() {
        assertTrue(ExcelDateUtil.isValidExcelDate(1));
        assertTrue(ExcelDateUtil.isValidExcelDate(45292));
        assertTrue(ExcelDateUtil.isValidExcelDate(73050));

        assertFalse(ExcelDateUtil.isValidExcelDate(0));
        assertFalse(ExcelDateUtil.isValidExcelDate(-1));
        assertFalse(ExcelDateUtil.isValidExcelDate(100000));
    }
}
