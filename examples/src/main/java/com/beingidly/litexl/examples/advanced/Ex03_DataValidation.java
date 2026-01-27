package com.beingidly.litexl.examples.advanced;

import com.beingidly.litexl.CellRange;
import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.format.DataValidation;
import com.beingidly.litexl.style.Style;

import java.nio.file.Path;

/**
 * Example: Data validation rules.
 *
 * This example demonstrates:
 * - Dropdown list validation
 * - Whole number range validation
 * - Decimal number validation
 * - Text length validation
 * - Custom formula validation
 */
public class Ex03_DataValidation {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex03_data_validation.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Validation");

            // Header style
            int headerStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .build());

            // === Dropdown List ===
            sheet.cell(0, 0).set("Status (dropdown)").style(headerStyle);
            sheet.cell(1, 0).set("");  // Empty cell for user input

            // Add list validation with dropdown
            sheet.addValidation(DataValidation.list(
                CellRange.of(1, 0),
                "Pending", "In Progress", "Completed", "Cancelled"
            ));

            // === Whole Number Range ===
            sheet.cell(0, 1).set("Age (1-120)").style(headerStyle);
            sheet.cell(1, 1).set("");

            sheet.addValidation(DataValidation.wholeNumber(
                CellRange.of(1, 1),
                1, 120
            ));

            // === Decimal Range ===
            sheet.cell(0, 2).set("Rating (0.0-5.0)").style(headerStyle);
            sheet.cell(1, 2).set("");

            sheet.addValidation(DataValidation.decimal(
                CellRange.of(1, 2),
                0.0, 5.0
            ));

            // === Text Length ===
            sheet.cell(0, 3).set("Code (3-10 chars)").style(headerStyle);
            sheet.cell(1, 3).set("");

            sheet.addValidation(DataValidation.textLength(
                CellRange.of(1, 3),
                3, 10
            ));

            // === Custom Formula ===
            sheet.cell(0, 4).set("Even number only").style(headerStyle);
            sheet.cell(1, 4).set("");

            sheet.addValidation(DataValidation.custom(
                CellRange.of(1, 4),
                "MOD(E2,2)=0",
                "Please enter an even number"
            ));

            // === Applying validation to multiple cells ===
            sheet.cell(3, 0).set("Department (multiple cells)").style(headerStyle);

            for (int i = 4; i <= 8; i++) {
                sheet.cell(i, 0).set("");
            }

            // Apply dropdown to range A5:A9
            sheet.addValidation(DataValidation.list(
                CellRange.of(4, 0, 8, 0),
                "Engineering", "Marketing", "Sales", "HR", "Finance"
            ));

            // Set column widths
            sheet.setColumnWidth(0, 25);
            sheet.setColumnWidth(1, 15);
            sheet.setColumnWidth(2, 18);
            sheet.setColumnWidth(3, 18);
            sheet.setColumnWidth(4, 18);

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
        System.out.println();
        System.out.println("Try entering invalid values in Excel to see validation in action:");
        System.out.println("- Column A: Click to see dropdown list");
        System.out.println("- Column B: Enter numbers outside 1-120 range");
        System.out.println("- Column C: Enter decimals outside 0.0-5.0");
        System.out.println("- Column D: Enter text shorter than 3 or longer than 10 characters");
        System.out.println("- Column E: Enter an odd number");
    }
}
