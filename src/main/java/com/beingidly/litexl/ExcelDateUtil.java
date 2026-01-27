package com.beingidly.litexl;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.temporal.ChronoUnit;

/**
 * Utility class for converting between Java dates and Excel serial dates.
 *
 * Excel uses a serial date system where:
 * - Day 1 = January 1, 1900 (but Excel incorrectly treats 1900 as a leap year)
 * - Time is represented as a fraction of a day
 *
 * Note: Excel has a bug where it considers 1900 to be a leap year (Feb 29, 1900 = day 60).
 * This class handles that bug for compatibility.
 */
final class ExcelDateUtil {

    private ExcelDateUtil() {}

    // Excel's epoch: December 31, 1899 (day 1 = January 1, 1900)
    private static final LocalDate EXCEL_EPOCH = LocalDate.of(1899, 12, 31);

    // The fake leap day in Excel (Feb 29, 1900 doesn't exist but Excel thinks it does)
    private static final int EXCEL_FAKE_LEAP_DAY = 60;

    // Seconds per day
    private static final double SECONDS_PER_DAY = 86400.0;

    /**
     * Converts a LocalDateTime to an Excel serial date number.
     */
    public static double toExcelDate(LocalDateTime dateTime) {
        LocalDate date = dateTime.toLocalDate();
        LocalTime time = dateTime.toLocalTime();

        double days = toExcelDate(date);
        double timeFraction = time.toSecondOfDay() / SECONDS_PER_DAY;

        return days + timeFraction;
    }

    /**
     * Converts a LocalDate to an Excel serial date number.
     */
    public static double toExcelDate(LocalDate date) {
        long days = ChronoUnit.DAYS.between(EXCEL_EPOCH, date);

        // Account for Excel's fake Feb 29, 1900
        // If date is after Feb 28, 1900, add 1 to account for the fake leap day
        if (days >= EXCEL_FAKE_LEAP_DAY) {
            days++;
        }

        return days;
    }

    /**
     * Converts an Excel serial date number to a LocalDateTime.
     */
    public static LocalDateTime fromExcelDate(double excelDate) {
        long days = (long) excelDate;
        double timeFraction = excelDate - days;

        // Account for Excel's fake Feb 29, 1900
        if (days >= EXCEL_FAKE_LEAP_DAY) {
            days--;
        }

        LocalDate date = EXCEL_EPOCH.plusDays(days);

        int totalSeconds = (int) Math.round(timeFraction * SECONDS_PER_DAY);
        int hours = totalSeconds / 3600;
        int minutes = (totalSeconds % 3600) / 60;
        int seconds = totalSeconds % 60;

        LocalTime time = LocalTime.of(hours, minutes, seconds);

        return LocalDateTime.of(date, time);
    }

    /**
     * Converts an Excel serial date number to a LocalDate (ignoring time).
     */
    public static LocalDate fromExcelDateToDate(double excelDate) {
        long days = (long) excelDate;

        // Account for Excel's fake Feb 29, 1900
        if (days >= EXCEL_FAKE_LEAP_DAY) {
            days--;
        }

        return EXCEL_EPOCH.plusDays(days);
    }

    /**
     * Checks if the given number appears to be an Excel date
     * (within reasonable range: 1900-2100).
     */
    public static boolean isValidExcelDate(double value) {
        // 1 = Jan 1, 1900
        // 73050 = Dec 31, 2099
        return value >= 1 && value <= 73050;
    }
}
