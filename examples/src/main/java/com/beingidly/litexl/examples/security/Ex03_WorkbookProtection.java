package com.beingidly.litexl.examples.security;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.crypto.WorkbookProtection;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;

/**
 * Example: Write protection and workbook structure protection.
 *
 * This example demonstrates:
 * - Setting write protection (recommends opening as read-only)
 * - Protecting workbook structure (prevents adding/deleting/renaming sheets)
 */
public class Ex03_WorkbookProtection {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex03_workbook_protection.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet1 = workbook.addSheet("Sales");
            sheet1.cell(0, 0).set("Product");
            sheet1.cell(0, 1).set("Revenue");
            sheet1.cell(1, 0).set("Widget A");
            sheet1.cell(1, 1).set(15000);

            Sheet sheet2 = workbook.addSheet("Summary");
            sheet2.cell(0, 0).set("Total Revenue");
            sheet2.cell(0, 1).setFormula("Sales!B2");

            // Set write protection - recommends opening as read-only
            workbook.setWriteProtection("modify123".toCharArray(), "Admin");

            // Protect workbook structure - prevents adding/deleting sheets
            workbook.protectStructure("struct456".toCharArray(), WorkbookProtection.defaults());

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
        System.out.println();
        System.out.println("Write protection password: 'modify123'");
        System.out.println("- Excel will recommend opening as read-only");
        System.out.println("- Enter the password to enable editing");
        System.out.println();
        System.out.println("Structure protection password: 'struct456'");
        System.out.println("- Cannot add, delete, or rename sheets");
    }
}
