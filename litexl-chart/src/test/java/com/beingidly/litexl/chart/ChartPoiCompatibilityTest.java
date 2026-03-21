package com.beingidly.litexl.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Cross-validation tests: write with LiteXL, read/verify with Apache POI.
 */
class ChartPoiCompatibilityTest {

    @TempDir
    Path tempDir;

    @Test
    void barChartReadableByPoi() throws Exception {
        Path file = tempDir.resolve("bar_poi.xlsx");

        // Write with LiteXL
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sales");
            sheet.cell(0, 0).set("Month");
            sheet.cell(0, 1).set("Revenue");
            sheet.cell(1, 0).set("Jan");
            sheet.cell(1, 1).set(1200);
            sheet.cell(2, 0).set("Feb");
            sheet.cell(2, 1).set(1500);
            sheet.cell(3, 0).set("Mar");
            sheet.cell(3, 1).set(1800);

            Chart chart = Chart.bar()
                .title("Monthly Revenue")
                .position("D1:K15")
                .barDirection(BarDirection.COLUMN)
                .grouping(Grouping.CLUSTERED)
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Sales!$A$2:$A$4")
                    .values("Sales!$B$2:$B$4")
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(file);
        }

        // Verify with Apache POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);
            assertEquals("Sales", sheet.getSheetName());

            // Verify cell data
            assertEquals("Jan", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(1200.0, sheet.getRow(1).getCell(1).getNumericCellValue());

            // Verify chart exists
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertNotNull(drawing, "Sheet should have a drawing");

            List<XSSFChart> charts = drawing.getCharts();
            assertEquals(1, charts.size(), "Should have 1 chart");

            XSSFChart chart = charts.getFirst();
            assertNotNull(chart, "Chart should not be null");
        }
    }

    @Test
    void lineChartReadableByPoi() throws Exception {
        Path file = tempDir.resolve("line_poi.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Trends");
            for (int i = 0; i < 5; i++) {
                sheet.cell(i, 0).set("Q" + (i + 1));
                sheet.cell(i, 1).set((i + 1) * 100.0);
            }

            Chart chart = Chart.line()
                .title("Quarterly Trend")
                .position("C1:J12")
                .addSeries(ChartSeries.of("Sales", "Trends!$A$1:$A$5", "Trends!$B$1:$B$5"))
                .build();

            Charts.add(sheet, chart);
            wb.save(file);
        }

        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertNotNull(drawing);

            List<XSSFChart> charts = drawing.getCharts();
            assertEquals(1, charts.size());
        }
    }

    @Test
    void pieChartReadableByPoi() throws Exception {
        Path file = tempDir.resolve("pie_poi.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Market");
            sheet.cell(0, 0).set("Product A");
            sheet.cell(0, 1).set(45);
            sheet.cell(1, 0).set("Product B");
            sheet.cell(1, 1).set(30);
            sheet.cell(2, 0).set("Product C");
            sheet.cell(2, 1).set(25);

            Chart chart = Chart.pie()
                .title("Market Share")
                .position("C1:J12")
                .addSeries(ChartSeries.of("Share", "Market!$A$1:$A$3", "Market!$B$1:$B$3"))
                .build();

            Charts.add(sheet, chart);
            wb.save(file);
        }

        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertNotNull(drawing);
            assertEquals(1, drawing.getCharts().size());
        }
    }

    @Test
    void multipleChartsReadableByPoi() throws Exception {
        Path file = tempDir.resolve("multi_poi.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            for (int i = 0; i < 5; i++) {
                sheet.cell(i, 0).set("Item " + i);
                sheet.cell(i, 1).set(i * 10.0);
                sheet.cell(i, 2).set(i * 5.0);
            }

            Charts.add(sheet, Chart.of(ChartType.BAR, "Chart 1",
                ChartPosition.of("D1:K10"),
                List.of(ChartSeries.of("S1", "Data!$A$1:$A$5", "Data!$B$1:$B$5"))));

            Charts.add(sheet, Chart.of(ChartType.LINE, "Chart 2",
                ChartPosition.of("D12:K22"),
                List.of(ChartSeries.of("S2", "Data!$A$1:$A$5", "Data!$C$1:$C$5"))));

            wb.save(file);
        }

        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertNotNull(drawing);
            assertEquals(2, drawing.getCharts().size());
        }
    }

    @Test
    void chartsOnMultipleSheets() throws Exception {
        Path file = tempDir.resolve("multi_sheets_poi.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet1 = wb.addSheet("Sheet1");
            sheet1.cell(0, 0).set("A");
            sheet1.cell(0, 1).set(10);
            Charts.add(sheet1, Chart.of(ChartType.BAR, "Chart A",
                ChartPosition.of("C1:J10"),
                List.of(ChartSeries.of("S", "Sheet1!$A$1", "Sheet1!$B$1"))));

            Sheet sheet2 = wb.addSheet("Sheet2");
            sheet2.cell(0, 0).set("B");
            sheet2.cell(0, 1).set(20);
            Charts.add(sheet2, Chart.of(ChartType.PIE, "Chart B",
                ChartPosition.of("C1:J10"),
                List.of(ChartSeries.of("S", "Sheet2!$A$1", "Sheet2!$B$1"))));

            wb.save(file);
        }

        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            // Sheet 1
            XSSFSheet sheet1 = wb.getSheetAt(0);
            assertNotNull(sheet1.getDrawingPatriarch());
            assertEquals(1, sheet1.getDrawingPatriarch().getCharts().size());

            // Sheet 2
            XSSFSheet sheet2 = wb.getSheetAt(1);
            assertNotNull(sheet2.getDrawingPatriarch());
            assertEquals(1, sheet2.getDrawingPatriarch().getCharts().size());
        }
    }
}
