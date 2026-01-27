package com.beingidly.litexl.examples.core;

import com.beingidly.litexl.CellValue;
import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;
import java.time.LocalDateTime;

/**
 * Example: Working with different cell types.
 *
 * This example demonstrates:
 * - String values
 * - Numeric values
 * - Boolean values
 * - Date/time values
 * - Reading values with type-safe pattern matching
 */
public class Ex03_CellTypes {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex03_cell_types.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("CellTypes");

            // Headers
            sheet.cell(0, 0).set("Type");
            sheet.cell(0, 1).set("Value");

            // String
            sheet.cell(1, 0).set("String");
            sheet.cell(1, 1).set("Hello, World!");

            // Number (integer)
            sheet.cell(2, 0).set("Integer");
            sheet.cell(2, 1).set(42);

            // Number (decimal)
            sheet.cell(3, 0).set("Decimal");
            sheet.cell(3, 1).set(3.14159);

            // Boolean
            sheet.cell(4, 0).set("Boolean");
            sheet.cell(4, 1).set(true);

            // Date/Time
            sheet.cell(5, 0).set("DateTime");
            sheet.cell(5, 1).set(LocalDateTime.of(2024, 12, 25, 10, 30));

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);

        // Demonstrate reading with pattern matching
        System.out.println();
        System.out.println("Reading back with pattern matching:");

        try (Workbook workbook = Workbook.open(outputPath)) {
            Sheet sheet = workbook.getSheet(0);

            for (int row = 1; row <= 5; row++) {
                CellValue value = sheet.cell(row, 1).value();

                String description = switch (value) {
                    case CellValue.Empty _ -> "Empty cell";
                    case CellValue.Text(String text) -> "Text: " + text;
                    case CellValue.Number(double num) -> "Number: " + num;
                    case CellValue.Bool(boolean bool) -> "Boolean: " + bool;
                    case CellValue.Date(var date) -> "Date: " + date;
                    case CellValue.Formula(String expr, CellValue cached) ->
                        "Formula: " + expr + " = " + cached;
                    case CellValue.Error(String code) -> "Error: " + code;
                };

                System.out.println("  Row " + row + ": " + description);
            }
        }
    }
}
