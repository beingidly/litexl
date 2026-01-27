package com.beingidly.litexl.examples.core;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;

/**
 * Example: Creating a new workbook and saving it.
 *
 * This example demonstrates:
 * - Creating a new workbook
 * - Adding a sheet
 * - Writing values to cells
 * - Saving to a file
 */
public class Ex01_CreateWorkbook {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex01_create_workbook.xlsx");

        // Create a new workbook using try-with-resources
        try (Workbook workbook = Workbook.create()) {

            // Add a sheet named "Hello"
            Sheet sheet = workbook.addSheet("Hello");

            // Write some values to cells (0-based indexing)
            sheet.cell(0, 0).set("Hello");
            sheet.cell(0, 1).set("World");
            sheet.cell(1, 0).set("litexl");
            sheet.cell(1, 1).set("is awesome!");

            // Save to file
            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
