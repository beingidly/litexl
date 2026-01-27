package com.beingidly.litexl.examples.core;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.style.Style;

import java.nio.file.Path;

/**
 * Example: Basic font styling.
 *
 * This example demonstrates:
 * - Bold text
 * - Italic text
 * - Font color
 * - Font size
 * - Underline
 * - Combining multiple styles
 */
public class Ex04_BasicStyling {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex04_basic_styling.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Styles");

            // Bold style
            int boldStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .build());

            // Italic style
            int italicStyle = workbook.addStyle(Style.builder()
                .italic(true)
                .build());

            // Red color style
            int redStyle = workbook.addStyle(Style.builder()
                .color(0xFFFF0000)  // ARGB: full alpha, red
                .build());

            // Large font style
            int largeStyle = workbook.addStyle(Style.builder()
                .font("Arial", 18)
                .build());

            // Underline style
            int underlineStyle = workbook.addStyle(Style.builder()
                .underline(true)
                .build());

            // Combined style: bold + italic + blue
            int combinedStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .italic(true)
                .color(0xFF0000FF)  // ARGB: blue
                .build());

            // Apply styles
            sheet.cell(0, 0).set("Normal text");

            sheet.cell(1, 0).set("Bold text").style(boldStyle);

            sheet.cell(2, 0).set("Italic text").style(italicStyle);

            sheet.cell(3, 0).set("Red text").style(redStyle);

            sheet.cell(4, 0).set("Large text (18pt)").style(largeStyle);

            sheet.cell(5, 0).set("Underlined text").style(underlineStyle);

            sheet.cell(6, 0).set("Bold + Italic + Blue").style(combinedStyle);

            // Set column width to fit content
            sheet.setColumnWidth(0, 25);

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
