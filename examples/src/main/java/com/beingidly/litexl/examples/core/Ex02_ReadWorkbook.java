package com.beingidly.litexl.examples.core;

import com.beingidly.litexl.Cell;
import com.beingidly.litexl.Row;
import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;

/**
 * Example: Reading an existing workbook.
 *
 * This example demonstrates:
 * - Opening an existing workbook
 * - Iterating over sheets
 * - Reading cell values
 */
public class Ex02_ReadWorkbook {

    public static void main(String[] args) {
        // First, create a sample file to read
        Path samplePath = createSampleFile();

        System.out.println("Reading workbook: " + samplePath);
        System.out.println();

        // Open the workbook
        try (Workbook workbook = Workbook.open(samplePath)) {

            // Print basic info
            System.out.println("Number of sheets: " + workbook.sheetCount());
            System.out.println();

            // Iterate over all sheets
            for (Sheet sheet : workbook.sheets()) {
                System.out.println("Sheet: " + sheet.name());
                System.out.println("  Rows: " + sheet.rowCount());

                // Iterate over rows
                for (var entry : sheet.rows().entrySet()) {
                    int rowIndex = entry.getKey();
                    Row row = entry.getValue();

                    System.out.println("  Row " + rowIndex + ":");

                    // Iterate over cells in the row
                    for (Cell cell : row.cells().values()) {
                        ExampleUtils.printCell(rowIndex, cell);
                    }
                }
                System.out.println();
            }
        }
    }

    private static Path createSampleFile() {
        Path path = ExampleUtils.tempFile("ex02_sample.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("SampleData");

            // Header row
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Age");
            sheet.cell(0, 2).set("City");

            // Data rows
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(30);
            sheet.cell(1, 2).set("Seoul");

            sheet.cell(2, 0).set("Bob");
            sheet.cell(2, 1).set(25);
            sheet.cell(2, 2).set("Tokyo");

            workbook.save(path);
        }

        return path;
    }
}
