package com.beingidly.litexl.examples.advanced;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.style.BorderStyle;
import com.beingidly.litexl.style.Style;

import java.nio.file.Path;

/**
 * Example: AutoFilter for data filtering.
 *
 * This example demonstrates:
 * - Setting up auto filter on a data range
 * - Creating a filterable table
 */
public class Ex05_AutoFilter {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex05_autofilter.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("AutoFilter");

            // Header style
            int headerStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .fill(0xFF4472C4)
                .color(0xFFFFFFFF)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            // Data style
            int dataStyle = workbook.addStyle(Style.builder()
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            // Headers
            String[] headers = {"ID", "Name", "Department", "Salary", "Status"};
            for (int i = 0; i < headers.length; i++) {
                sheet.cell(0, i).set(headers[i]).style(headerStyle);
            }

            // Sample employee data
            Object[][] data = {
                {1, "Alice", "Engineering", 95000, "Active"},
                {2, "Bob", "Marketing", 75000, "Active"},
                {3, "Charlie", "Engineering", 85000, "Active"},
                {4, "Diana", "Sales", 70000, "Inactive"},
                {5, "Eve", "Engineering", 105000, "Active"},
                {6, "Frank", "HR", 65000, "Active"},
                {7, "Grace", "Marketing", 80000, "Active"},
                {8, "Henry", "Sales", 72000, "Inactive"},
                {9, "Ivy", "Engineering", 92000, "Active"},
                {10, "Jack", "HR", 68000, "Active"}
            };

            for (int row = 0; row < data.length; row++) {
                for (int col = 0; col < data[row].length; col++) {
                    Object value = data[row][col];
                    if (value instanceof Integer) {
                        sheet.cell(row + 1, col).set((int) value).style(dataStyle);
                    } else if (value instanceof String) {
                        sheet.cell(row + 1, col).set((String) value).style(dataStyle);
                    }
                }
            }

            // Set auto filter on the data range (including header)
            // Range: A1:E11 (row 0 to row 10, col 0 to col 4)
            sheet.setAutoFilter(0, 0, data.length, headers.length - 1);

            // Set column widths
            sheet.setColumnWidth(0, 5);
            sheet.setColumnWidth(1, 12);
            sheet.setColumnWidth(2, 15);
            sheet.setColumnWidth(3, 12);
            sheet.setColumnWidth(4, 10);

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
        System.out.println();
        System.out.println("Open in Excel to see filter dropdowns in the header row.");
        System.out.println("Try filtering by:");
        System.out.println("- Department: Engineering, Marketing, Sales, HR");
        System.out.println("- Status: Active, Inactive");
        System.out.println("- Salary: Sort ascending/descending or filter by value");
    }
}
