package com.beingidly.litexl.examples.util;

import com.beingidly.litexl.Cell;
import com.beingidly.litexl.CellValue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Utility methods shared across litexl examples.
 */
public final class ExampleUtils {

    private ExampleUtils() {
        // Utility class
    }

    /**
     * Creates a temporary file path with the given name.
     * The file is created in the system temp directory.
     *
     * @param name the file name (e.g., "example.xlsx")
     * @return path to the temporary file
     */
    public static Path tempFile(String name) {
        try {
            Path tempDir = Files.createTempDirectory("litexl-examples");
            return tempDir.resolve(name);
        } catch (IOException e) {
            throw new RuntimeException("Failed to create temp directory", e);
        }
    }

    /**
     * Prints a message indicating a file was created.
     *
     * @param path the path to the created file
     */
    public static void printCreated(Path path) {
        System.out.println("Created: " + path.toAbsolutePath());
        System.out.println("Open this file in Excel or LibreOffice Calc to view the result.");
    }

    /**
     * Prints the value of a cell with type information.
     *
     * @param row the row index
     * @param cell the cell to print
     */
    public static void printCell(int row, Cell cell) {
        CellValue value = cell.value();
        String typeStr = switch (value) {
            case CellValue.Empty _ -> "Empty";
            case CellValue.Text(String text) -> "Text: \"" + text + "\"";
            case CellValue.Number(double num) -> "Number: " + num;
            case CellValue.Bool(boolean bool) -> "Bool: " + bool;
            case CellValue.Date(var date) -> "Date: " + date;
            case CellValue.Formula(String expr, CellValue cached) -> "Formula: " + expr + " (cached: " + cached + ")";
            case CellValue.Error(String code) -> "Error: " + code;
        };
        System.out.println("  Cell[" + row + "," + cell.column() + "]: " + typeStr);
    }

    /**
     * Prints a section header for better readability.
     *
     * @param title the section title
     */
    public static void printSection(String title) {
        System.out.println();
        System.out.println("=== " + title + " ===");
    }
}
