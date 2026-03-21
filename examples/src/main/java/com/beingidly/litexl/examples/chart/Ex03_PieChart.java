package com.beingidly.litexl.examples.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.style.*;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;

/**
 * Example: Pie and doughnut charts.
 *
 * This example demonstrates:
 * - Pie chart with data labels showing percentages
 * - Pie chart with exploded slices
 * - Custom fill colors per series
 * - Doughnut chart
 * - Legend positioning
 */
public class Ex03_PieChart {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex03_pie_chart.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Pie Charts");

            // === Write sample data ===
            sheet.cell(0, 0).set("Browser");
            sheet.cell(0, 1).set("Market Share %");
            sheet.cell(1, 0).set("Chrome");    sheet.cell(1, 1).set(64.7);
            sheet.cell(2, 0).set("Safari");    sheet.cell(2, 1).set(18.6);
            sheet.cell(3, 0).set("Firefox");   sheet.cell(3, 1).set(3.3);
            sheet.cell(4, 0).set("Edge");      sheet.cell(4, 1).set(5.2);
            sheet.cell(5, 0).set("Others");    sheet.cell(5, 1).set(8.2);

            sheet.setColumnWidth(0, 12);
            sheet.setColumnWidth(1, 16);

            // === Pie chart with percentage labels ===
            Chart pieChart = Chart.pie()
                .title("Browser Market Share")
                .position("C1:K16")
                .addSeries(ChartSeries.builder()
                    .name("Market Share")
                    .categories("$A$2:$A$6")
                    .values("$B$2:$B$6")
                    .dataLabel(ChartDataLabel.builder()
                        .showPercent(true)
                        .showCategory(true)
                        .showValue(false)
                        .separator("\n")
                        .build())
                    .build())
                .legend(LegendPosition.RIGHT)
                .build();

            Charts.add(sheet, pieChart);

            // === Pie chart with exploded slice ===
            Chart explodedPie = Chart.pie()
                .title("Exploded Pie")
                .position("C18:K33")
                .addSeries(ChartSeries.builder()
                    .name("Share")
                    .categories("$A$2:$A$6")
                    .values("$B$2:$B$6")
                    .explosion(15)
                    .dataLabel(ChartDataLabel.withPercent())
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, explodedPie);

            // === Doughnut chart ===
            Chart doughnut = Chart.builder(ChartType.DOUGHNUT)
                .title("Doughnut Chart")
                .position("L1:T16")
                .addSeries(ChartSeries.builder()
                    .name("Share")
                    .categories("$A$2:$A$6")
                    .values("$B$2:$B$6")
                    .dataLabel(ChartDataLabel.builder()
                        .showValue(true)
                        .showCategory(false)
                        .build())
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, doughnut);

            wb.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
