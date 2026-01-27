package com.beingidly.litexl.examples.core;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.style.Style;

import java.nio.file.Path;
import java.time.LocalDateTime;

/**
 * Example: Number and date formats.
 *
 * This example demonstrates:
 * - Number formats (thousands separator, decimals)
 * - Currency formats
 * - Percentage formats
 * - Date formats
 * - Time formats
 */
public class Ex07_NumberFormat {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex07_number_format.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("NumberFormats");

            // Number with thousands separator
            int thousandsStyle = workbook.addStyle(Style.builder()
                .format("#,##0")
                .build());

            // Number with decimals
            int decimalsStyle = workbook.addStyle(Style.builder()
                .format("#,##0.00")
                .build());

            // Currency (USD)
            int currencyStyle = workbook.addStyle(Style.builder()
                .format("$#,##0.00")
                .build());

            // Percentage
            int percentStyle = workbook.addStyle(Style.builder()
                .format("0.0%")
                .build());

            // Date format (YYYY-MM-DD)
            int dateStyle = workbook.addStyle(Style.builder()
                .format("yyyy-mm-dd")
                .build());

            // Date format (Month Day, Year)
            int dateLongStyle = workbook.addStyle(Style.builder()
                .format("mmmm d, yyyy")
                .build());

            // Time format
            int timeStyle = workbook.addStyle(Style.builder()
                .format("hh:mm:ss AM/PM")
                .build());

            // DateTime format
            int dateTimeStyle = workbook.addStyle(Style.builder()
                .format("yyyy-mm-dd hh:mm")
                .build());

            // Header style
            int headerStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .build());

            // Headers
            sheet.cell(0, 0).set("Format").style(headerStyle);
            sheet.cell(0, 1).set("Value").style(headerStyle);
            sheet.cell(0, 2).set("Formatted").style(headerStyle);

            // Thousands separator
            sheet.cell(1, 0).set("Thousands (#,##0)");
            sheet.cell(1, 1).set(1234567);
            sheet.cell(1, 2).set(1234567).style(thousandsStyle);

            // Decimals
            sheet.cell(2, 0).set("Decimals (#,##0.00)");
            sheet.cell(2, 1).set(1234.5);
            sheet.cell(2, 2).set(1234.5).style(decimalsStyle);

            // Currency
            sheet.cell(3, 0).set("Currency ($#,##0.00)");
            sheet.cell(3, 1).set(9999.99);
            sheet.cell(3, 2).set(9999.99).style(currencyStyle);

            // Percentage (value is stored as decimal, e.g., 0.75 = 75%)
            sheet.cell(4, 0).set("Percentage (0.0%)");
            sheet.cell(4, 1).set(0.756);
            sheet.cell(4, 2).set(0.756).style(percentStyle);

            // Date
            LocalDateTime date = LocalDateTime.of(2024, 12, 25, 0, 0);
            sheet.cell(5, 0).set("Date (yyyy-mm-dd)");
            sheet.cell(5, 1).set(date);
            sheet.cell(5, 2).set(date).style(dateStyle);

            // Long date
            sheet.cell(6, 0).set("Long date (mmmm d, yyyy)");
            sheet.cell(6, 1).set(date);
            sheet.cell(6, 2).set(date).style(dateLongStyle);

            // Time
            LocalDateTime time = LocalDateTime.of(2024, 1, 1, 14, 30, 45);
            sheet.cell(7, 0).set("Time (hh:mm:ss AM/PM)");
            sheet.cell(7, 1).set(time);
            sheet.cell(7, 2).set(time).style(timeStyle);

            // DateTime
            LocalDateTime dateTime = LocalDateTime.of(2024, 12, 25, 10, 30);
            sheet.cell(8, 0).set("DateTime (yyyy-mm-dd hh:mm)");
            sheet.cell(8, 1).set(dateTime);
            sheet.cell(8, 2).set(dateTime).style(dateTimeStyle);

            // Set column widths
            sheet.setColumnWidth(0, 30);
            sheet.setColumnWidth(1, 20);
            sheet.setColumnWidth(2, 25);

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
