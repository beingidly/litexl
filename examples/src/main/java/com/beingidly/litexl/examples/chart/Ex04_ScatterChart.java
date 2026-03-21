package com.beingidly.litexl.examples.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.axis.*;
import com.beingidly.litexl.chart.style.*;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;

/**
 * Example: Scatter (XY) charts.
 *
 * This example demonstrates:
 * - Scatter chart with markers only
 * - Scatter chart with smooth lines and markers
 * - Different scatter styles (MARKER, LINE_MARKER, SMOOTH_MARKER)
 * - Custom marker styles and sizes
 * - X/Y axis configuration
 */
public class Ex04_ScatterChart {

    /** Example runner. */
    private Ex04_ScatterChart() {}

    /**
     * Runs the example.
     *
     * @param args command-line arguments (not used)
     */
    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex04_scatter_chart.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Scatter Charts");

            // === Write sample data ===
            sheet.cell(0, 0).set("X");
            sheet.cell(0, 1).set("Y (Linear)");
            sheet.cell(0, 2).set("Y (Quadratic)");
            for (int i = 0; i < 15; i++) {
                double x = i * 1.0;
                sheet.cell(i + 1, 0).set(x);
                sheet.cell(i + 1, 1).set(2.5 * x + 3);
                sheet.cell(i + 1, 2).set(0.3 * x * x + 1);
            }

            // === Scatter with markers ===
            Chart scatterMarkers = Chart.scatter()
                .title("Data Distribution (Markers)")
                .position("D1:M16")
                .scatterStyle(ScatterStyle.MARKER)
                .addSeries(ChartSeries.builder()
                    .name("Linear")
                    .categories("$A$2:$A$16")
                    .values("$B$2:$B$16")
                    .marker(ChartMarker.of(MarkerStyle.CIRCLE, 8))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Quadratic")
                    .categories("$A$2:$A$16")
                    .values("$C$2:$C$16")
                    .marker(ChartMarker.of(MarkerStyle.TRIANGLE, 7))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, scatterMarkers);

            // === Scatter with smooth lines + markers ===
            Chart scatterSmooth = Chart.scatter()
                .title("Smooth Curve Fit")
                .position("D18:M34")
                .scatterStyle(ScatterStyle.SMOOTH_MARKER)
                .addSeries(ChartSeries.builder()
                    .name("Quadratic")
                    .categories("$A$2:$A$16")
                    .values("$C$2:$C$16")
                    .line(ChartLine.builder()
                        .color(ChartColor.rgb("FF6347"))
                        .width(2.0)
                        .build())
                    .marker(ChartMarker.of(MarkerStyle.DIAMOND, 6))
                    .smooth(true)
                    .build())
                .legend(LegendPosition.NONE)
                .valueAxis(ValueAxis.builder(1, 2)
                    .title("X Values")
                    .position(AxisPosition.BOTTOM)
                    .build())
                .valueAxis(ValueAxis.builder(2, 1)
                    .title("Y Values")
                    .position(AxisPosition.LEFT)
                    .numberFormat("#,##0.0")
                    .build())
                .build();

            Charts.add(sheet, scatterSmooth);

            wb.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
