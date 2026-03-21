package com.beingidly.litexl.examples.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.style.ChartColor;
import com.beingidly.litexl.chart.style.ChartFill;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;
import java.util.List;

/**
 * Example: Bar and column charts.
 *
 * This example demonstrates:
 * - Creating a simple bar chart with static factory
 * - Creating a column chart with builder
 * - Clustered vs stacked grouping
 * - Custom bar direction
 * - Multiple series with custom fill colors
 */
public class Ex01_BarChart {

    /** Example runner. */
    private Ex01_BarChart() {}

    /**
     * Runs the example.
     *
     * @param args command-line arguments (not used)
     */
    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex01_bar_chart.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Bar Charts");

            // === Write sample data ===
            String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun"};
            double[] revenue = {12000, 15000, 13500, 18000, 16500, 21000};
            double[] costs = {8000, 9500, 8800, 11000, 10200, 12500};

            sheet.cell(0, 0).set("Month");
            sheet.cell(0, 1).set("Revenue");
            sheet.cell(0, 2).set("Costs");
            for (int i = 0; i < months.length; i++) {
                sheet.cell(i + 1, 0).set(months[i]);
                sheet.cell(i + 1, 1).set(revenue[i]);
                sheet.cell(i + 1, 2).set(costs[i]);
            }

            // Set column widths
            sheet.setColumnWidth(0, 10);
            sheet.setColumnWidth(1, 12);
            sheet.setColumnWidth(2, 12);

            // === Chart 1: Simple bar chart (static factory) ===
            Chart simpleBar = Chart.of(
                ChartType.BAR,
                "Revenue by Month (Horizontal Bar)",
                ChartPosition.of("D1:L12"),
                List.of(ChartSeries.of("Revenue", "$A$2:$A$7", "$B$2:$B$7"))
            );
            Charts.add(sheet, simpleBar);

            // === Chart 2: Column chart with multiple series (builder) ===
            Chart columnChart = Chart.column()
                .title("Revenue vs Costs (Clustered Column)")
                .position("D14:L26")
                .grouping(Grouping.CLUSTERED)
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("$A$2:$A$7")
                    .values("$B$2:$B$7")
                    .fill(ChartFill.solid(ChartColor.rgb("4472C4")))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("$A$2:$A$7")
                    .values("$C$2:$C$7")
                    .fill(ChartFill.solid(ChartColor.rgb("ED7D31")))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();
            Charts.add(sheet, columnChart);

            // === Chart 3: Stacked bar chart ===
            Chart stackedBar = Chart.bar()
                .title("Revenue + Costs (Stacked)")
                .position("D28:L40")
                .grouping(Grouping.STACKED)
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("$A$2:$A$7")
                    .values("$B$2:$B$7")
                    .fill(ChartFill.solid(ChartColor.rgb("5B9BD5")))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("$A$2:$A$7")
                    .values("$C$2:$C$7")
                    .fill(ChartFill.solid(ChartColor.rgb("FFC000")))
                    .build())
                .legend(LegendPosition.RIGHT)
                .build();
            Charts.add(sheet, stackedBar);

            wb.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
