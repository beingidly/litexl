package com.beingidly.litexl.examples.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.axis.*;
import com.beingidly.litexl.chart.style.*;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;

/**
 * Example: Line charts.
 *
 * This example demonstrates:
 * - Line chart with markers
 * - Smooth line option
 * - Custom line color, width, and dash style
 * - Custom axis configuration (title, number format, min/max)
 * - Multiple series with different marker styles
 */
public class Ex02_LineChart {

    /** Example runner. */
    private Ex02_LineChart() {}

    /**
     * Runs the example.
     *
     * @param args command-line arguments (not used)
     */
    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex02_line_chart.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Line Charts");

            // === Write sample data ===
            String[] quarters = {"Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25"};
            double[] actual = {42.5, 45.1, 48.3, 52.0, 55.8, 59.2};
            double[] target = {40.0, 44.0, 48.0, 52.0, 56.0, 60.0};

            sheet.cell(0, 0).set("Quarter");
            sheet.cell(0, 1).set("Actual");
            sheet.cell(0, 2).set("Target");
            for (int i = 0; i < quarters.length; i++) {
                sheet.cell(i + 1, 0).set(quarters[i]);
                sheet.cell(i + 1, 1).set(actual[i]);
                sheet.cell(i + 1, 2).set(target[i]);
            }

            // === Line chart with markers and custom axes ===
            Chart chart = Chart.line()
                .title("Actual vs Target Performance")
                .position("D1:N18")
                .addSeries(ChartSeries.builder()
                    .name("Actual")
                    .categories("$A$2:$A$7")
                    .values("$B$2:$B$7")
                    .line(ChartLine.builder()
                        .color(ChartColor.rgb("2E75B6"))
                        .width(2.5)
                        .dash(LineDash.SOLID)
                        .build())
                    .marker(ChartMarker.of(MarkerStyle.CIRCLE, 8))
                    .smooth(false)
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Target")
                    .categories("$A$2:$A$7")
                    .values("$C$2:$C$7")
                    .line(ChartLine.builder()
                        .color(ChartColor.rgb("A5A5A5"))
                        .width(1.5)
                        .dash(LineDash.DASH)
                        .build())
                    .marker(ChartMarker.of(MarkerStyle.DIAMOND, 6))
                    .smooth(false)
                    .build())
                .legend(LegendPosition.BOTTOM)
                .categoryAxis(CategoryAxis.builder(1, 2)
                    .title("Quarter")
                    .position(AxisPosition.BOTTOM)
                    .build())
                .valueAxis(ValueAxis.builder(2, 1)
                    .title("Revenue (M$)")
                    .position(AxisPosition.LEFT)
                    .numberFormat("#,##0.0")
                    .minimum(30.0)
                    .maximum(70.0)
                    .majorUnit(10.0)
                    .build())
                .build();

            Charts.add(sheet, chart);

            // === Smooth line chart ===
            Chart smoothChart = Chart.line()
                .title("Smooth Trend Line")
                .position("D20:N36")
                .addSeries(ChartSeries.builder()
                    .name("Actual")
                    .categories("$A$2:$A$7")
                    .values("$B$2:$B$7")
                    .line(ChartLine.of("FF6347", 3.0))
                    .marker(ChartMarker.of(MarkerStyle.SQUARE, 7))
                    .smooth(true)
                    .build())
                .legend(LegendPosition.NONE)
                .build();

            Charts.add(sheet, smoothChart);

            wb.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
