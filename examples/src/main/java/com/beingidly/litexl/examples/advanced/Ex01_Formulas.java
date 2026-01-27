package com.beingidly.litexl.examples.advanced;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.style.Style;

import java.nio.file.Path;

/**
 * Example: Using formulas in Excel.
 *
 * This example demonstrates:
 * - Basic arithmetic formulas
 * - SUM, AVERAGE functions
 * - Cell references
 * - Formula with cached values
 */
public class Ex01_Formulas {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex01_formulas.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Formulas");

            // Header style
            int headerStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .build());

            // === Basic Arithmetic ===
            sheet.cell(0, 0).set("Basic Arithmetic").style(headerStyle);

            sheet.cell(1, 0).set("A");
            sheet.cell(1, 1).set(10);

            sheet.cell(2, 0).set("B");
            sheet.cell(2, 1).set(5);

            sheet.cell(3, 0).set("A + B");
            sheet.cell(3, 1).setFormula("B2+B3");

            sheet.cell(4, 0).set("A * B");
            sheet.cell(4, 1).setFormula("B2*B3");

            sheet.cell(5, 0).set("A / B");
            sheet.cell(5, 1).setFormula("B2/B3");

            // === SUM Function ===
            sheet.cell(0, 3).set("SUM Function").style(headerStyle);

            sheet.cell(1, 3).set("Value 1");
            sheet.cell(1, 4).set(100);

            sheet.cell(2, 3).set("Value 2");
            sheet.cell(2, 4).set(200);

            sheet.cell(3, 3).set("Value 3");
            sheet.cell(3, 4).set(300);

            sheet.cell(4, 3).set("Total");
            sheet.cell(4, 4).setFormula("SUM(E2:E4)");

            // === AVERAGE Function ===
            sheet.cell(0, 6).set("AVERAGE Function").style(headerStyle);

            sheet.cell(1, 6).set("Score 1");
            sheet.cell(1, 7).set(85);

            sheet.cell(2, 6).set("Score 2");
            sheet.cell(2, 7).set(92);

            sheet.cell(3, 6).set("Score 3");
            sheet.cell(3, 7).set(78);

            sheet.cell(4, 6).set("Average");
            sheet.cell(4, 7).setFormula("AVERAGE(H2:H4)");

            // === Sales Table with Formulas ===
            sheet.cell(7, 0).set("Sales Report").style(headerStyle);

            sheet.cell(8, 0).set("Product").style(headerStyle);
            sheet.cell(8, 1).set("Qty").style(headerStyle);
            sheet.cell(8, 2).set("Price").style(headerStyle);
            sheet.cell(8, 3).set("Subtotal").style(headerStyle);

            String[] products = {"Laptop", "Mouse", "Keyboard"};
            int[] quantities = {5, 20, 10};
            double[] prices = {999.99, 29.99, 79.99};

            for (int i = 0; i < products.length; i++) {
                int row = 9 + i;
                sheet.cell(row, 0).set(products[i]);
                sheet.cell(row, 1).set(quantities[i]);
                sheet.cell(row, 2).set(prices[i]);
                sheet.cell(row, 3).setFormula("B" + (row + 1) + "*C" + (row + 1));
            }

            // Grand total
            sheet.cell(12, 2).set("Grand Total:").style(headerStyle);
            sheet.cell(12, 3).setFormula("SUM(D10:D12)");

            // Set column widths
            sheet.setColumnWidth(0, 15);
            sheet.setColumnWidth(3, 12);
            sheet.setColumnWidth(6, 12);

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
        System.out.println();
        System.out.println("Note: Formulas will be calculated when opened in Excel.");
    }
}
