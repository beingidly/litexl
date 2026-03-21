package com.beingidly.litexl.examples.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.style.ChartColor;
import com.beingidly.litexl.chart.style.ChartFill;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;

/**
 * Example: Area and radar charts.
 *
 * This example demonstrates:
 * - Area chart with stacked grouping
 * - Radar chart
 * - Different grouping modes (STANDARD, STACKED, PERCENT_STACKED)
 */
public class Ex05_AreaChart {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex05_area_chart.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Area Charts");

            // === Write sample data ===
            String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun"};
            double[] webTraffic = {1200, 1500, 1800, 1600, 2100, 2400};
            double[] mobileTraffic = {800, 1100, 1300, 1500, 1700, 2000};
            double[] apiTraffic = {400, 500, 600, 700, 800, 1000};

            sheet.cell(0, 0).set("Month");
            sheet.cell(0, 1).set("Web");
            sheet.cell(0, 2).set("Mobile");
            sheet.cell(0, 3).set("API");
            for (int i = 0; i < months.length; i++) {
                sheet.cell(i + 1, 0).set(months[i]);
                sheet.cell(i + 1, 1).set(webTraffic[i]);
                sheet.cell(i + 1, 2).set(mobileTraffic[i]);
                sheet.cell(i + 1, 3).set(apiTraffic[i]);
            }

            // === Stacked area chart ===
            Chart stackedArea = Chart.area()
                .title("Traffic by Channel (Stacked Area)")
                .position("E1:O18")
                .grouping(Grouping.STACKED)
                .addSeries(ChartSeries.builder()
                    .name("Web")
                    .categories("$A$2:$A$7")
                    .values("$B$2:$B$7")
                    .fill(ChartFill.solid(ChartColor.rgb("4472C4")))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Mobile")
                    .categories("$A$2:$A$7")
                    .values("$C$2:$C$7")
                    .fill(ChartFill.solid(ChartColor.rgb("ED7D31")))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("API")
                    .categories("$A$2:$A$7")
                    .values("$D$2:$D$7")
                    .fill(ChartFill.solid(ChartColor.rgb("A5A5A5")))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, stackedArea);

            // === Radar chart ===
            // Write radar data
            String[] skills = {"Java", "SQL", "DevOps", "Design", "Testing", "Communication"};
            double[] teamA = {9, 7, 6, 4, 8, 7};
            double[] teamB = {6, 8, 8, 7, 5, 9};

            int radarRow = 9;
            sheet.cell(radarRow, 0).set("Skill");
            sheet.cell(radarRow, 1).set("Team A");
            sheet.cell(radarRow, 2).set("Team B");
            for (int i = 0; i < skills.length; i++) {
                sheet.cell(radarRow + 1 + i, 0).set(skills[i]);
                sheet.cell(radarRow + 1 + i, 1).set(teamA[i]);
                sheet.cell(radarRow + 1 + i, 2).set(teamB[i]);
            }

            Chart radar = Chart.builder(ChartType.RADAR)
                .title("Team Skills Comparison")
                .position("E20:O36")
                .radarStyle(RadarStyle.MARKER)
                .addSeries(ChartSeries.builder()
                    .name("Team A")
                    .categories("$A$11:$A$16")
                    .values("$B$11:$B$16")
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Team B")
                    .categories("$A$11:$A$16")
                    .values("$C$11:$C$16")
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, radar);

            wb.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
