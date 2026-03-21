package com.beingidly.litexl.examples.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.style.ChartColor;
import com.beingidly.litexl.chart.style.ChartFill;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;

/**
 * Example: 3D charts.
 *
 * This example demonstrates:
 * - 3D bar chart with custom 3D view rotation and perspective
 * - 3D pie chart
 * - 3D area chart
 * - ChartView3D configuration (xRotation, yRotation, perspective)
 */
public class Ex07_3DChart {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex07_3d_chart.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("3D Charts");

            // === Write sample data ===
            String[] regions = {"North", "South", "East", "West"};
            double[] sales = {45000, 38000, 52000, 41000};
            double[] growth = {12, 8, 15, 10};

            sheet.cell(0, 0).set("Region");
            sheet.cell(0, 1).set("Sales");
            sheet.cell(0, 2).set("Growth %");
            for (int i = 0; i < regions.length; i++) {
                sheet.cell(i + 1, 0).set(regions[i]);
                sheet.cell(i + 1, 1).set(sales[i]);
                sheet.cell(i + 1, 2).set(growth[i]);
            }

            // === 3D Bar chart with custom view ===
            Chart bar3d = Chart.builder(ChartType.BAR_3D)
                .title("Regional Sales (3D)")
                .position("D1:N16")
                .view3D(ChartView3D.of(15, 20, 30))
                .grouping(Grouping.CLUSTERED)
                .addSeries(ChartSeries.builder()
                    .name("Sales")
                    .categories("$A$2:$A$5")
                    .values("$B$2:$B$5")
                    .fill(ChartFill.solid(ChartColor.rgb("4472C4")))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, bar3d);

            // === 3D Pie chart ===
            Chart pie3d = Chart.builder(ChartType.PIE_3D)
                .title("Sales Distribution (3D Pie)")
                .position("D18:N33")
                .view3D(ChartView3D.of(30, 0, 0))
                .addSeries(ChartSeries.builder()
                    .name("Sales")
                    .categories("$A$2:$A$5")
                    .values("$B$2:$B$5")
                    .explosion(10)
                    .build())
                .legend(LegendPosition.RIGHT)
                .build();

            Charts.add(sheet, pie3d);

            // === 3D Area chart ===
            Chart area3d = Chart.builder(ChartType.AREA_3D)
                .title("Growth Trends (3D Area)")
                .position("O1:Y16")
                .view3D(ChartView3D.of(20, 30, 25))
                .grouping(Grouping.STANDARD)
                .addSeries(ChartSeries.builder()
                    .name("Growth %")
                    .categories("$A$2:$A$5")
                    .values("$C$2:$C$5")
                    .fill(ChartFill.gradient(
                        ChartColor.rgb("70AD47"),
                        ChartColor.rgb("2E7D32")))
                    .build())
                .legend(LegendPosition.NONE)
                .build();

            Charts.add(sheet, area3d);

            wb.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
