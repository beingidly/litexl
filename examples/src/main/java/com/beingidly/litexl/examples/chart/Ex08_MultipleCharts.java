package com.beingidly.litexl.examples.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.style.ChartColor;
import com.beingidly.litexl.chart.style.ChartDataLabel;
import com.beingidly.litexl.chart.style.ChartFill;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;
import java.util.List;

/**
 * Example: Multiple charts on one sheet and across sheets.
 *
 * This example demonstrates:
 * - Multiple charts on a single sheet
 * - Charts across multiple sheets
 * - Using Charts.add() to attach charts to sheets
 * - Chart without title
 * - Using literal data (NumberArray) instead of cell references
 */
public class Ex08_MultipleCharts {

    /** Example runner. */
    private Ex08_MultipleCharts() {}

    /**
     * Runs the example.
     *
     * @param args command-line arguments (not used)
     */
    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex08_multiple_charts.xlsx");

        try (Workbook wb = Workbook.create()) {
            // === Sheet 1: Dashboard with multiple charts ===
            Sheet dashboard = wb.addSheet("Dashboard");

            String[] products = {"Laptop", "Phone", "Tablet", "Watch", "Earbuds"};
            double[] q1 = {120, 250, 80, 45, 180};
            double[] q2 = {135, 280, 75, 60, 210};

            dashboard.cell(0, 0).set("Product");
            dashboard.cell(0, 1).set("Q1 Sales");
            dashboard.cell(0, 2).set("Q2 Sales");
            for (int i = 0; i < products.length; i++) {
                dashboard.cell(i + 1, 0).set(products[i]);
                dashboard.cell(i + 1, 1).set(q1[i]);
                dashboard.cell(i + 1, 2).set(q2[i]);
            }

            // Chart 1: Bar chart (top-left)
            Charts.add(dashboard, Chart.column()
                .title("Q1 vs Q2 Sales")
                .position("D1:L14")
                .grouping(Grouping.CLUSTERED)
                .addSeries(ChartSeries.builder()
                    .name("Q1")
                    .categories("$A$2:$A$6")
                    .values("$B$2:$B$6")
                    .fill(ChartFill.solid(ChartColor.rgb("4472C4")))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Q2")
                    .categories("$A$2:$A$6")
                    .values("$C$2:$C$6")
                    .fill(ChartFill.solid(ChartColor.rgb("ED7D31")))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build());

            // Chart 2: Pie chart (top-right)
            Charts.add(dashboard, Chart.pie()
                .title("Q2 Market Mix")
                .position("M1:U14")
                .addSeries(ChartSeries.builder()
                    .name("Q2")
                    .categories("$A$2:$A$6")
                    .values("$C$2:$C$6")
                    .dataLabel(ChartDataLabel.withPercent())
                    .build())
                .legend(LegendPosition.RIGHT)
                .build());

            // Chart 3: Line chart (bottom) — without title
            Charts.add(dashboard, Chart.of(
                ChartType.LINE,
                ChartPosition.of("D16:L28"),
                List.of(
                    ChartSeries.of("Q1", "$A$2:$A$6", "$B$2:$B$6"),
                    ChartSeries.of("Q2", "$A$2:$A$6", "$C$2:$C$6")
                )));

            // === Sheet 2: Summary with its own chart ===
            Sheet summary = wb.addSheet("Summary");

            summary.cell(0, 0).set("Quarter");
            summary.cell(0, 1).set("Total");
            summary.cell(1, 0).set("Q1");
            summary.cell(1, 1).set(675);
            summary.cell(2, 0).set("Q2");
            summary.cell(2, 1).set(760);

            Charts.add(summary, Chart.bar()
                .title("Quarterly Summary")
                .position("C1:J12")
                .addSeries(ChartSeries.builder()
                    .name("Total Sales")
                    .categories("$A$2:$A$3")
                    .values("$B$2:$B$3")
                    .fill(ChartFill.solid(ChartColor.rgb("70AD47")))
                    .build())
                .legend(LegendPosition.NONE)
                .build());

            // === Sheet 3: Chart with literal data ===
            Sheet literal = wb.addSheet("Literal Data");

            Charts.add(literal, Chart.pie()
                .title("Inline Data Chart")
                .position("A1:H14")
                .addSeries(ChartSeries.builder()
                    .name("Ratings")
                    .categories(ChartDataSource.ofStrings("Excellent", "Good", "Average", "Poor"))
                    .values(ChartDataSource.ofNumbers(40, 35, 15, 10))
                    .dataLabel(ChartDataLabel.builder()
                        .showValue(true)
                        .showCategory(true)
                        .separator(" - ")
                        .build())
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build());

            wb.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
