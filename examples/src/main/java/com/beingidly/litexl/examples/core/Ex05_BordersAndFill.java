package com.beingidly.litexl.examples.core;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.style.BorderStyle;
import com.beingidly.litexl.style.Style;

import java.nio.file.Path;

/**
 * Example: Borders and fill colors.
 *
 * This example demonstrates:
 * - Thin borders on all sides
 * - Different border styles
 * - Background fill colors
 * - Combining borders with fill
 */
public class Ex05_BordersAndFill {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex05_borders_fill.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("BordersAndFill");

            // Thin border on all sides
            int thinBorderStyle = workbook.addStyle(Style.builder()
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            // Medium border
            int mediumBorderStyle = workbook.addStyle(Style.builder()
                .border(BorderStyle.MEDIUM, 0xFF000000)
                .build());

            // Thick border
            int thickBorderStyle = workbook.addStyle(Style.builder()
                .border(BorderStyle.THICK, 0xFF000000)
                .build());

            // Dashed border
            int dashedBorderStyle = workbook.addStyle(Style.builder()
                .border(BorderStyle.DASHED, 0xFF000000)
                .build());

            // Yellow fill
            int yellowFillStyle = workbook.addStyle(Style.builder()
                .fill(0xFFFFFF00)  // ARGB: yellow
                .build());

            // Blue fill
            int blueFillStyle = workbook.addStyle(Style.builder()
                .fill(0xFF4472C4)  // ARGB: Excel blue
                .color(0xFFFFFFFF) // White text
                .build());

            // Green fill with border
            int greenWithBorderStyle = workbook.addStyle(Style.builder()
                .fill(0xFF70AD47)  // ARGB: Excel green
                .color(0xFFFFFFFF) // White text
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            // Bottom border only
            int bottomBorderStyle = workbook.addStyle(Style.builder()
                .borderBottom(BorderStyle.MEDIUM, 0xFF000000)
                .build());

            // Apply styles - Row 0: Border styles
            sheet.cell(0, 0).set("Borders:").style(bottomBorderStyle);
            sheet.cell(1, 0).set("Thin border").style(thinBorderStyle);
            sheet.cell(2, 0).set("Medium border").style(mediumBorderStyle);
            sheet.cell(3, 0).set("Thick border").style(thickBorderStyle);
            sheet.cell(4, 0).set("Dashed border").style(dashedBorderStyle);

            // Column B: Fill styles
            sheet.cell(0, 1).set("Fills:").style(bottomBorderStyle);
            sheet.cell(1, 1).set("Yellow fill").style(yellowFillStyle);
            sheet.cell(2, 1).set("Blue fill").style(blueFillStyle);
            sheet.cell(3, 1).set("Green + border").style(greenWithBorderStyle);

            // Set column widths
            sheet.setColumnWidth(0, 20);
            sheet.setColumnWidth(1, 20);

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
