package com.beingidly.litexl.examples.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;
import java.util.List;

/**
 * Example: Reading and writing charts.
 *
 * This example demonstrates:
 * - Creating a workbook with charts and saving it
 * - Opening an existing workbook and reading charts back
 * - Inspecting chart properties (type, title, series count)
 * - Sheet name auto-inference (omitting sheet name in references)
 */
public class Ex09_ChartReadWrite {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex09_chart_read_write.xlsx");

        // === Step 1: Create and save ===
        ExampleUtils.printSection("Creating Workbook with Charts");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sales Data");

            sheet.cell(0, 0).set("Product");
            sheet.cell(0, 1).set("Revenue");
            sheet.cell(0, 2).set("Profit");
            String[] products = {"Widget A", "Widget B", "Widget C", "Widget D"};
            double[] revenue = {50000, 35000, 28000, 42000};
            double[] profit = {15000, 10000, 8000, 12000};
            for (int i = 0; i < products.length; i++) {
                sheet.cell(i + 1, 0).set(products[i]);
                sheet.cell(i + 1, 1).set(revenue[i]);
                sheet.cell(i + 1, 2).set(profit[i]);
            }

            // Sheet name auto-inference: no "Sales Data!" prefix needed
            Chart chart = Chart.column()
                .title("Product Performance")
                .position("D1:L15")
                .grouping(Grouping.CLUSTERED)
                .addSeries(ChartSeries.of("Revenue", "$A$2:$A$5", "$B$2:$B$5"))
                .addSeries(ChartSeries.of("Profit", "$A$2:$A$5", "$C$2:$C$5"))
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);

            System.out.println("Created chart: " + chart.type() + " - " + chart.title().text());
            System.out.println("  Series count: " + chart.series().size());
            System.out.println("  Position: col " + chart.position().fromCol()
                + " to col " + chart.position().toCol());

            wb.save(outputPath);
            System.out.println("Saved to: " + outputPath);
        }

        // === Step 2: Read back ===
        ExampleUtils.printSection("Reading Charts from Saved File");

        try (Workbook wb = Workbook.open(outputPath)) {
            for (int i = 0; i < wb.sheetCount(); i++) {
                Sheet sheet = wb.getSheet(i);
                List<Chart> charts = Charts.get(sheet);

                System.out.println("Sheet: " + sheet.name() + " (" + charts.size() + " charts)");

                for (int j = 0; j < charts.size(); j++) {
                    Chart chart = charts.get(j);
                    System.out.println("  Chart " + (j + 1) + ":");
                    System.out.println("    Type: " + chart.type());
                    System.out.println("    Title: " + (chart.title() != null ? chart.title().text() : "(none)"));
                    System.out.println("    Series: " + chart.series().size());
                    for (ChartSeries s : chart.series()) {
                        System.out.println("      - " + (s.name() != null ? s.name() : "(unnamed)"));
                    }
                }
            }
        }

        ExampleUtils.printCreated(outputPath);
    }
}
