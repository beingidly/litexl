package com.beingidly.litexl.examples.advanced;

import com.beingidly.litexl.CellRange;
import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.format.ConditionalFormat;
import com.beingidly.litexl.style.Style;

import java.nio.file.Path;

/**
 * Example: Conditional formatting.
 *
 * This example demonstrates:
 * - Highlighting cells based on value comparison
 * - Highlighting duplicates
 * - Using formula-based conditional formatting
 */
public class Ex04_ConditionalFormat {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex04_conditional_format.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("ConditionalFormat");

            // Styles for conditional formatting
            int greenStyle = workbook.addStyle(Style.builder()
                .fill(0xFFC6EFCE)  // Light green
                .color(0xFF006100) // Dark green text
                .build());

            int redStyle = workbook.addStyle(Style.builder()
                .fill(0xFFFFC7CE)  // Light red
                .color(0xFF9C0006) // Dark red text
                .build());

            int yellowStyle = workbook.addStyle(Style.builder()
                .fill(0xFFFFEB9C)  // Light yellow
                .color(0xFF9C5700) // Dark yellow text
                .build());

            // Header style
            int headerStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .build());

            // === Value Comparison ===
            sheet.cell(0, 0).set("Score").style(headerStyle);
            sheet.cell(0, 1).set("Status").style(headerStyle);

            int[] scores = {95, 72, 88, 45, 60, 78, 92, 55};

            for (int i = 0; i < scores.length; i++) {
                sheet.cell(i + 1, 0).set(scores[i]);
            }

            // Green for scores >= 80
            sheet.addConditionalFormat(ConditionalFormat.greaterThan(
                CellRange.of(1, 0, 8, 0),
                79,
                greenStyle
            ));

            // Yellow for scores between 60 and 79
            sheet.addConditionalFormat(ConditionalFormat.between(
                CellRange.of(1, 0, 8, 0),
                60, 79,
                yellowStyle
            ));

            // Red for scores < 60
            sheet.addConditionalFormat(ConditionalFormat.lessThan(
                CellRange.of(1, 0, 8, 0),
                60,
                redStyle
            ));

            // === Duplicate Detection ===
            sheet.cell(0, 3).set("Names (duplicates highlighted)").style(headerStyle);

            String[] names = {"Alice", "Bob", "Charlie", "Alice", "David", "Bob", "Eve"};

            for (int i = 0; i < names.length; i++) {
                sheet.cell(i + 1, 3).set(names[i]);
            }

            // Highlight duplicates
            sheet.addConditionalFormat(ConditionalFormat.duplicateValues(
                CellRange.of(1, 3, 7, 3),
                redStyle
            ));

            // === Formula-based Conditional Format ===
            sheet.cell(0, 5).set("Value").style(headerStyle);
            sheet.cell(0, 6).set("Target").style(headerStyle);
            sheet.cell(0, 7).set("(Green if Value >= Target)").style(headerStyle);

            int[] values = {100, 80, 150, 90, 120};
            int[] targets = {90, 100, 100, 95, 110};

            for (int i = 0; i < values.length; i++) {
                sheet.cell(i + 1, 5).set(values[i]);
                sheet.cell(i + 1, 6).set(targets[i]);
            }

            // Highlight when value >= target using formula
            sheet.addConditionalFormat(ConditionalFormat.expression(
                CellRange.of(1, 5, 5, 5),
                "F2>=G2",
                greenStyle
            ));

            // Set column widths
            sheet.setColumnWidth(0, 10);
            sheet.setColumnWidth(3, 30);
            sheet.setColumnWidth(5, 10);
            sheet.setColumnWidth(6, 10);
            sheet.setColumnWidth(7, 30);

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
        System.out.println();
        System.out.println("Open in Excel to see conditional formatting:");
        System.out.println("- Column A: Scores colored by grade (green >= 80, yellow 60-79, red < 60)");
        System.out.println("- Column D: Duplicate names highlighted in red");
        System.out.println("- Column F: Values highlighted green when >= target");
    }
}
