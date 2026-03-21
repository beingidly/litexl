package com.beingidly.litexl.examples.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.style.*;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;
import java.util.List;

/**
 * Example: Chart styling.
 *
 * This example demonstrates:
 * - Solid fill colors (RGB, preset, theme)
 * - Gradient fill (two-stop and multi-stop)
 * - Pattern fill
 * - Custom line properties (width, dash, cap)
 * - Custom font (title font)
 * - Data labels with custom separator
 * - Error bars
 */
public class Ex06_StyledChart {

    /** Example runner. */
    private Ex06_StyledChart() {}

    /**
     * Runs the example.
     *
     * @param args command-line arguments (not used)
     */
    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex06_styled_chart.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Styled Charts");

            // === Write sample data ===
            String[] items = {"Alpha", "Beta", "Gamma", "Delta", "Epsilon"};
            double[] values = {85, 72, 93, 61, 78};

            sheet.cell(0, 0).set("Item");
            sheet.cell(0, 1).set("Score");
            for (int i = 0; i < items.length; i++) {
                sheet.cell(i + 1, 0).set(items[i]);
                sheet.cell(i + 1, 1).set(values[i]);
            }

            // === Chart with gradient fill ===
            Chart gradientChart = Chart.column()
                .title(ChartTitle.of("Gradient Fill", ChartFont.builder()
                    .name("Arial")
                    .size(14)
                    .bold(true)
                    .color(ChartColor.rgb("1F4E79"))
                    .build()))
                .position("C1:L16")
                .addSeries(ChartSeries.builder()
                    .name("Score")
                    .categories("$A$2:$A$6")
                    .values("$B$2:$B$6")
                    .fill(ChartFill.gradient(
                        ChartColor.rgb("4472C4"),
                        ChartColor.rgb("2F5597"),
                        90.0))
                    .line(ChartLine.builder()
                        .color(ChartColor.rgb("1F4E79"))
                        .width(1.5)
                        .build())
                    .dataLabel(ChartDataLabel.builder()
                        .showValue(true)
                        .build())
                    .build())
                .legend(LegendPosition.NONE)
                .build();

            Charts.add(sheet, gradientChart);

            // === Chart with pattern fill ===
            Chart patternChart = Chart.bar()
                .title("Pattern Fill Comparison")
                .position("C18:L33")
                .addSeries(ChartSeries.builder()
                    .name("Score")
                    .categories("$A$2:$A$6")
                    .values("$B$2:$B$6")
                    .fill(ChartFill.pattern(
                        PatternType.DIAGONAL_CROSS,
                        ChartColor.rgb("4472C4"),
                        ChartColor.rgb("D6E4F0")))
                    .build())
                .legend(LegendPosition.NONE)
                .build();

            Charts.add(sheet, patternChart);

            // === Chart with error bars ===
            Chart errorBarChart = Chart.column()
                .title("Scores with Error Bars (5%)")
                .position("M1:V16")
                .addSeries(ChartSeries.builder()
                    .name("Score")
                    .categories("$A$2:$A$6")
                    .values("$B$2:$B$6")
                    .fill(ChartFill.solid(ChartColor.rgb("70AD47")))
                    .errorBars(ChartErrorBars.percentage(5.0))
                    .build())
                .legend(LegendPosition.NONE)
                .build();

            Charts.add(sheet, errorBarChart);

            // === Chart with preset colors ===
            Chart presetChart = Chart.column()
                .title("Preset Colors")
                .position("M18:V33")
                .addSeries(ChartSeries.builder()
                    .name("Score")
                    .categories("$A$2:$A$6")
                    .values("$B$2:$B$6")
                    .fill(ChartFill.solid(ChartColor.preset(PresetColor.CORAL)))
                    .line(ChartLine.builder()
                        .color(ChartColor.preset(PresetColor.CRIMSON))
                        .width(2.0)
                        .dash(LineDash.DASH_DOT)
                        .build())
                    .build())
                .legend(LegendPosition.NONE)
                .build();

            Charts.add(sheet, presetChart);

            wb.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
