package com.beingidly.litexl.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.axis.*;
import com.beingidly.litexl.chart.style.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Integration tests: write charts with litexl, read back, verify round-trip.
 */
class ChartWriteReadTest {

    @TempDir
    Path tempDir;

    @Test
    void simpleBarChart() throws Exception {
        Path file = tempDir.resolve("bar.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Category");
            sheet.cell(0, 1).set("Value");
            sheet.cell(1, 0).set("A");
            sheet.cell(1, 1).set(10);
            sheet.cell(2, 0).set("B");
            sheet.cell(2, 1).set(20);
            sheet.cell(3, 0).set("C");
            sheet.cell(3, 1).set(30);

            Chart chart = Chart.of(ChartType.BAR, "Sales",
                ChartPosition.of("D1:K15"),
                List.of(ChartSeries.of("Revenue", "$A$2:$A$4", "$B$2:$B$4")));

            Charts.add(sheet, chart);
            assertEquals(1, Charts.get(sheet).size());
            wb.save(file);
        }

        // Read back
        try (Workbook wb = Workbook.open(file)) {
            Sheet sheet = wb.getSheet(0);
            assertNotNull(sheet);
            List<Chart> charts = Charts.get(sheet);
            assertEquals(1, charts.size());

            Chart chart = charts.getFirst();
            assertNotNull(chart.title());
            assertEquals("Sales", chart.title().text());
            assertEquals(1, chart.series().size());
        }
    }

    @Test
    void lineChartWithBuilder() throws Exception {
        Path file = tempDir.resolve("line.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            for (int i = 0; i < 10; i++) {
                sheet.cell(i, 0).set("Item " + i);
                sheet.cell(i, 1).set(i * 10.0);
                sheet.cell(i, 2).set(i * 5.0);
            }

            Chart chart = Chart.line()
                .title("Trends")
                .position("N1:V20")
                .addSeries(ChartSeries.builder()
                    .name("Series A")
                    .categories("$A$1:$A$10")
                    .values("$B$1:$B$10")
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Series B")
                    .categories("$A$1:$A$10")
                    .values("$C$1:$C$10")
                    .fill(ChartFill.solid("FF0000"))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            List<Chart> charts = Charts.get(wb.getSheet(0));
            assertEquals(1, charts.size());
            assertEquals(ChartType.LINE, charts.getFirst().type());
        }
    }

    @Test
    void pieChart() throws Exception {
        Path file = tempDir.resolve("pie.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Market");
            sheet.cell(0, 0).set("Product");
            sheet.cell(0, 1).set("Share");
            sheet.cell(1, 0).set("Alpha");
            sheet.cell(1, 1).set(40);
            sheet.cell(2, 0).set("Beta");
            sheet.cell(2, 1).set(35);
            sheet.cell(3, 0).set("Gamma");
            sheet.cell(3, 1).set(25);

            Chart chart = Chart.pie()
                .title("Market Share")
                .position("D1:K15")
                .addSeries(ChartSeries.builder()
                    .name("Share")
                    .categories("$A$2:$A$4")
                    .values("$B$2:$B$4")
                    .dataLabel(ChartDataLabel.withPercent())
                    .build())
                .build();

            Charts.add(sheet, chart);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            List<Chart> charts = Charts.get(wb.getSheet(0));
            assertEquals(1, charts.size());
            assertEquals(ChartType.PIE, charts.getFirst().type());
        }
    }

    @Test
    void multipleChartsOnOneSheet() throws Exception {
        Path file = tempDir.resolve("multi.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            for (int i = 0; i < 5; i++) {
                sheet.cell(i, 0).set("Item " + i);
                sheet.cell(i, 1).set(i * 10.0);
            }

            Charts.add(sheet, Chart.of(ChartType.BAR, "Chart 1",
                ChartPosition.of("C1:J10"),
                List.of(ChartSeries.of("S1", "$A$1:$A$5", "$B$1:$B$5"))));

            Charts.add(sheet, Chart.of(ChartType.LINE, "Chart 2",
                ChartPosition.of("C12:J22"),
                List.of(ChartSeries.of("S2", "$A$1:$A$5", "$B$1:$B$5"))));

            assertEquals(2, Charts.get(sheet).size());
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            List<Chart> charts = Charts.get(wb.getSheet(0));
            assertEquals(2, charts.size());
        }
    }

    @Test
    void chartWithAutoInferredSheetName() throws Exception {
        Path file = tempDir.resolve("auto_infer.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("My Sheet");
            sheet.cell(0, 0).set("A");
            sheet.cell(0, 1).set(1);

            // No sheet name in reference - should auto-infer "My Sheet"
            Chart chart = Chart.of(ChartType.BAR, "Test",
                ChartPosition.of("C1:J10"),
                List.of(ChartSeries.of("S", "$A$1:$A$1", "$B$1:$B$1")));

            Charts.add(sheet, chart);
            wb.save(file);
        }

        // If it saved without error, auto-inference worked
        try (Workbook wb = Workbook.open(file)) {
            assertNotNull(wb.getSheet(0));
            assertEquals(1, Charts.get(wb.getSheet(0)).size());
        }
    }

    @Test
    void scatterChart() throws Exception {
        Path file = tempDir.resolve("scatter.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("XY");
            for (int i = 0; i < 10; i++) {
                sheet.cell(i, 0).set(i * 1.0);
                sheet.cell(i, 1).set(i * i * 1.0);
            }

            Chart chart = Chart.scatter()
                .title("XY Plot")
                .position("C1:K15")
                .scatterStyle(ScatterStyle.LINE_MARKER)
                .addSeries(ChartSeries.builder()
                    .name("Quadratic")
                    .categories("$A$1:$A$10")
                    .values("$B$1:$B$10")
                    .marker(MarkerStyle.CIRCLE, 5)
                    .build())
                .build();

            Charts.add(sheet, chart);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            List<Chart> charts = Charts.get(wb.getSheet(0));
            assertEquals(1, charts.size());
            assertEquals(ChartType.SCATTER, charts.getFirst().type());
        }
    }

    @Test
    void chartsStaticApi() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.cell(0, 0).set(1);

            assertEquals(0, Charts.get(sheet).size());

            Chart chart = Chart.of(ChartType.BAR, ChartPosition.of("A1:D5"),
                List.of(ChartSeries.of("$A$1")));
            Charts.add(sheet, chart);
            assertEquals(1, Charts.get(sheet).size());

            Charts.add(sheet, chart);
            assertEquals(2, Charts.get(sheet).size());

            Charts.remove(sheet, 0);
            assertEquals(1, Charts.get(sheet).size());

            Charts.clear(sheet);
            assertEquals(0, Charts.get(sheet).size());
        }
    }
}
