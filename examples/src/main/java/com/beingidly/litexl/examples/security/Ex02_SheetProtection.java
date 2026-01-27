package com.beingidly.litexl.examples.security;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.crypto.SheetProtection;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.style.Style;

import java.nio.file.Path;

/**
 * Example: Protecting sheets with password.
 *
 * This example demonstrates:
 * - Protecting a sheet with password
 * - Configuring which actions are allowed
 * - Making certain cells editable while protecting others
 */
public class Ex02_SheetProtection {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex02_sheet_protection.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Protected");

            // Create styles: locked (default) and unlocked
            int lockedStyle = workbook.addStyle(Style.builder()
                .locked(true)
                .fill(0xFFD9D9D9)  // Gray background for locked cells
                .build());

            int unlockedStyle = workbook.addStyle(Style.builder()
                .locked(false)
                .fill(0xFFFFFFCC)  // Yellow background for editable cells
                .build());

            // Header row (locked)
            sheet.cell(0, 0).set("Item").style(lockedStyle);
            sheet.cell(0, 1).set("Quantity (editable)").style(lockedStyle);
            sheet.cell(0, 2).set("Price").style(lockedStyle);
            sheet.cell(0, 3).set("Total").style(lockedStyle);

            // Data rows
            String[] items = {"Apple", "Banana", "Orange"};
            double[] prices = {1.50, 0.75, 2.00};

            for (int i = 0; i < items.length; i++) {
                int row = i + 1;
                sheet.cell(row, 0).set(items[i]).style(lockedStyle);
                sheet.cell(row, 1).set(0).style(unlockedStyle);  // Editable quantity
                sheet.cell(row, 2).set(prices[i]).style(lockedStyle);
                sheet.cell(row, 3).setFormula("B" + (row + 1) + "*C" + (row + 1)).style(lockedStyle);
            }

            // Set column widths
            sheet.setColumnWidth(0, 15);
            sheet.setColumnWidth(1, 20);
            sheet.setColumnWidth(2, 10);
            sheet.setColumnWidth(3, 10);

            // Protect the sheet with custom permissions
            SheetProtection protection = SheetProtection.builder()
                .selectLockedCells(true)      // Can select locked cells
                .selectUnlockedCells(true)    // Can select unlocked cells
                .formatCells(false)           // Cannot format cells
                .insertRows(false)            // Cannot insert rows
                .deleteRows(false)            // Cannot delete rows
                .sort(true)                   // Can sort data
                .autoFilter(true)             // Can use auto filter
                .build();

            // Apply protection with password
            sheet.protectionManager().protect("editme".toCharArray(), protection);

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
        System.out.println();
        System.out.println("The sheet is protected with password: 'editme'");
        System.out.println("- Gray cells are locked (read-only)");
        System.out.println("- Yellow cells (Quantity column) are editable");
        System.out.println("- Total column uses formulas that update automatically");
    }
}
