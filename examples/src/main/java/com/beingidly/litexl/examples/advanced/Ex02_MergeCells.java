package com.beingidly.litexl.examples.advanced;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.style.BorderStyle;
import com.beingidly.litexl.style.HAlign;
import com.beingidly.litexl.style.Style;
import com.beingidly.litexl.style.VAlign;

import java.nio.file.Path;

/**
 * Example: Merging cells.
 *
 * This example demonstrates:
 * - Merging cells horizontally
 * - Merging cells vertically
 * - Creating a table with merged header
 */
public class Ex02_MergeCells {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex02_merge_cells.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("MergedCells");

            // Title style - centered in merged area
            int titleStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .font("Arial", 16)
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .fill(0xFF4472C4)
                .color(0xFFFFFFFF)
                .build());

            // Header style
            int headerStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .fill(0xFFD9E1F2)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            // Data style
            int dataStyle = workbook.addStyle(Style.builder()
                .border(BorderStyle.THIN, 0xFF000000)
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .build());

            // === Horizontal merge - Title spanning columns A-E ===
            sheet.cell(0, 0).set("Quarterly Sales Report 2024").style(titleStyle);
            sheet.merge(0, 0, 0, 4);  // Merge row 0, columns 0-4
            sheet.row(0).height(30);

            // === Table with merged headers ===
            // Main header row with merged cells
            sheet.cell(2, 0).set("Region").style(headerStyle);
            sheet.merge(2, 0, 3, 0);  // Vertical merge for "Region"

            sheet.cell(2, 1).set("Q1-Q2").style(headerStyle);
            sheet.merge(2, 1, 2, 2);  // Horizontal merge for "Q1-Q2"

            sheet.cell(2, 3).set("Q3-Q4").style(headerStyle);
            sheet.merge(2, 3, 2, 4);  // Horizontal merge for "Q3-Q4"

            // Sub-headers
            sheet.cell(3, 1).set("Q1").style(headerStyle);
            sheet.cell(3, 2).set("Q2").style(headerStyle);
            sheet.cell(3, 3).set("Q3").style(headerStyle);
            sheet.cell(3, 4).set("Q4").style(headerStyle);

            // Data rows
            String[] regions = {"North", "South", "East", "West"};
            int[][] sales = {
                {120, 145, 160, 180},
                {90, 110, 125, 140},
                {150, 165, 175, 190},
                {80, 95, 105, 115}
            };

            for (int i = 0; i < regions.length; i++) {
                int row = 4 + i;
                sheet.cell(row, 0).set(regions[i]).style(dataStyle);
                for (int j = 0; j < 4; j++) {
                    sheet.cell(row, 1 + j).set(sales[i][j]).style(dataStyle);
                }
            }

            // Set row heights for merged cells
            sheet.row(2).height(25);
            sheet.row(3).height(25);

            // Set column widths
            sheet.setColumnWidth(0, 12);
            for (int i = 1; i <= 4; i++) {
                sheet.setColumnWidth(i, 10);
            }

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
