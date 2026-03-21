package com.beingidly.litexl.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.chart.axis.*;
import com.beingidly.litexl.chart.style.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Comparator;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive comparison: creates identical charts with LiteXL and Apache POI,
 * compares the internal XML structure, and verifies cross-read compatibility.
 *
 * <p>Files are also saved to /tmp/chart-comparison/ for manual inspection in Excel.
 */
class ChartComprehensiveComparisonTest {

    @TempDir
    Path tempDir;

    static final Path OUTPUT_DIR = Path.of("/tmp/chart-comparison");

    // === Shared test data ===
    static final String[] CATEGORIES = {"Jan", "Feb", "Mar", "Apr", "May", "Jun"};
    static final double[] REVENUE = {12000, 15000, 13500, 18000, 16500, 21000};
    static final double[] COSTS = {8000, 9500, 8800, 11000, 10200, 12500};
    static final double[] PROFIT = {4000, 5500, 4700, 7000, 6300, 8500};

    static final String[] PIE_LABELS = {"Product A", "Product B", "Product C", "Product D", "Product E"};
    static final double[] PIE_VALUES = {35, 25, 20, 15, 5};

    // ===========================
    // 1. Clustered Column Chart
    // ===========================
    @Test
    void compareClusteredColumn() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_clustered_column.xlsx");
        Path poiFile = tempDir.resolve("poi_clustered_column.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.column()
                .title("Revenue vs Costs")
                .position("E1:L15")
                .grouping(Grouping.CLUSTERED)
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.solid("4472C4"))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .fill(ChartFill.solid("ED7D31"))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Revenue vs Costs");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var costData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));

            XDDFBarChartData barData = (XDDFBarChartData) chart.createData(ChartTypes.BAR, catAxis, valAxis);
            barData.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.COL);

            var s1 = barData.addSeries(cats, revData);
            s1.setTitle("Revenue", null);
            var s2 = barData.addSeries(cats, costData);
            s2.setTitle("Costs", null);

            chart.plot(barData);

            // Solid fills via CT API
            var ctBar = chart.getCTChart().getPlotArea().getBarChartArray(0);
            ctBar.getSerArray(0).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            ctBar.getSerArray(1).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "clustered_column");
    }

    // ===========================
    // 2. Stacked Bar Chart
    // ===========================
    @Test
    void compareStackedBar() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_stacked_bar.xlsx");
        Path poiFile = tempDir.resolve("poi_stacked_bar.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.bar()
                .title("Stacked Revenue Breakdown")
                .position("E1:L15")
                .grouping(Grouping.STACKED)
                .barDirection(BarDirection.BAR)
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.solid("4472C4"))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .fill(ChartFill.solid("ED7D31"))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Profit")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$D$2:$D$7")
                    .fill(ChartFill.solid("A5A5A5"))
                    .build())
                .legend(LegendPosition.RIGHT)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Stacked Revenue Breakdown");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var costData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));
            var profitData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 3, 3));

            XDDFBarChartData barData = (XDDFBarChartData) chart.createData(ChartTypes.BAR, catAxis, valAxis);
            barData.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.BAR);
            barData.setBarGrouping(BarGrouping.STACKED);

            barData.addSeries(cats, revData).setTitle("Revenue", null);
            barData.addSeries(cats, costData).setTitle("Costs", null);
            barData.addSeries(cats, profitData).setTitle("Profit", null);

            chart.plot(barData);

            var ctBar = chart.getCTChart().getPlotArea().getBarChartArray(0);
            ctBar.getSerArray(0).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            ctBar.getSerArray(1).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});
            ctBar.getSerArray(2).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{(byte) 0xA5, (byte) 0xA5, (byte) 0xA5});

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.RIGHT);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "stacked_bar");
    }

    // ===========================
    // 3. Line Chart with Markers
    // ===========================
    @Test
    void compareLineWithMarkers() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_line_markers.xlsx");
        Path poiFile = tempDir.resolve("poi_line_markers.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.line()
                .title("Trend Analysis")
                .position("E1:L15")
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .line(ChartLine.builder()
                        .color("4472C4")
                        .width(2.5)
                        .dash(LineDash.SOLID)
                        .build())
                    .marker(MarkerStyle.CIRCLE, 6)
                    .smooth(true)
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .line(ChartLine.builder()
                        .color("ED7D31")
                        .width(2.0)
                        .dash(LineDash.DASH)
                        .build())
                    .marker(MarkerStyle.SQUARE, 5)
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Trend Analysis");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var costData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));

            XDDFLineChartData lineData = (XDDFLineChartData) chart.createData(ChartTypes.LINE, catAxis, valAxis);

            var s1 = (XDDFLineChartData.Series) lineData.addSeries(cats, revData);
            s1.setTitle("Revenue", null);
            s1.setSmooth(true);
            s1.setMarkerStyle(org.apache.poi.xddf.usermodel.chart.MarkerStyle.CIRCLE);
            s1.setMarkerSize((short) 6);

            var s2 = (XDDFLineChartData.Series) lineData.addSeries(cats, costData);
            s2.setTitle("Costs", null);
            s2.setSmooth(false);
            s2.setMarkerStyle(org.apache.poi.xddf.usermodel.chart.MarkerStyle.SQUARE);
            s2.setMarkerSize((short) 5);

            chart.plot(lineData);

            // Line styles via CT API
            var ctLine = chart.getCTChart().getPlotArea().getLineChartArray(0);
            var ln1 = ctLine.getSerArray(0).addNewSpPr().addNewLn();
            ln1.setW(31750); // 2.5pt * 12700
            ln1.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});

            var ln2 = ctLine.getSerArray(1).addNewSpPr().addNewLn();
            ln2.setW(25400); // 2.0pt * 12700
            ln2.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});
            ln2.addNewPrstDash().setVal(org.openxmlformats.schemas.drawingml.x2006.main.STPresetLineDashVal.DASH);

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "line_markers");
    }

    // ===========================
    // 4. Pie Chart with Labels
    // ===========================
    @Test
    void comparePieWithLabels() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_pie.xlsx");
        Path poiFile = tempDir.resolve("poi_pie.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Market");
            writePieData(sheet);

            Chart chart = Chart.pie()
                .title("Market Share")
                .position("D1:K15")
                .addSeries(ChartSeries.builder()
                    .name("Share")
                    .categories("Market!$A$2:$A$6")
                    .values("Market!$B$2:$B$6")
                    .dataLabel(ChartDataLabel.builder()
                        .showPercent(true)
                        .showCategory(true)
                        .separator("\n")
                        .showLeaderLines(true)
                        .build())
                    .explosion(10)
                    .build())
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Market");
            writePoiPieData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 0, 10, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Market Share");

            // POI pie charts: create dummy axes (ignored by pie)
            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 5, 0, 0));
            var vals = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 5, 1, 1));

            XDDFChartData pieData = chart.createData(ChartTypes.PIE, catAxis, valAxis);
            var s1 = pieData.addSeries(cats, vals);
            s1.setTitle("Share", null);

            chart.plot(pieData);

            // Data labels + explosion via CT API
            var ctPie = chart.getCTChart().getPlotArea().getPieChartArray(0);
            var ctSer = ctPie.getSerArray(0);

            // Explosion
            ctSer.addNewExplosion().setVal(10);

            // Data labels
            var dLbls = ctSer.addNewDLbls();
            dLbls.addNewShowVal().setVal(false);
            dLbls.addNewShowCatName().setVal(true);
            dLbls.addNewShowSerName().setVal(false);
            dLbls.addNewShowPercent().setVal(true);
            dLbls.addNewShowLegendKey().setVal(false);
            dLbls.setSeparator("\n");
            dLbls.addNewShowLeaderLines().setVal(true);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "pie");
    }

    // ===========================
    // 5. Scatter Chart
    // ===========================
    @Test
    void compareScatter() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_scatter.xlsx");
        Path poiFile = tempDir.resolve("poi_scatter.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("XY");
            sheet.cell(0, 0).set("X");
            sheet.cell(0, 1).set("Y1");
            sheet.cell(0, 2).set("Y2");
            for (int i = 0; i < 10; i++) {
                sheet.cell(i + 1, 0).set(i * 1.0);
                sheet.cell(i + 1, 1).set(i * i * 1.0);
                sheet.cell(i + 1, 2).set(i * 2.0);
            }

            Chart chart = Chart.scatter()
                .title("XY Scatter Plot")
                .position("E1:L15")
                .scatterStyle(ScatterStyle.LINE_MARKER)
                .addSeries(ChartSeries.builder()
                    .name("Quadratic")
                    .categories("XY!$A$2:$A$11")
                    .values("XY!$B$2:$B$11")
                    .marker(MarkerStyle.CIRCLE, 5)
                    .line(ChartLine.of("4472C4", 1.5))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Linear")
                    .categories("XY!$A$2:$A$11")
                    .values("XY!$C$2:$C$11")
                    .marker(MarkerStyle.DIAMOND, 5)
                    .line(ChartLine.of("ED7D31", 1.5))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("XY");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("X");
            header.createCell(1).setCellValue("Y1");
            header.createCell(2).setCellValue("Y2");
            for (int i = 0; i < 10; i++) {
                var row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(i * 1.0);
                row.createCell(1).setCellValue(i * i * 1.0);
                row.createCell(2).setCellValue(i * 2.0);
            }

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("XY Scatter Plot");

            XDDFValueAxis xAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis yAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            yAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var xData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 10, 0, 0));
            var y1Data = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 10, 1, 1));
            var y2Data = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 10, 2, 2));

            XDDFScatterChartData scatterData = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, xAxis, yAxis);
            scatterData.setStyle(org.apache.poi.xddf.usermodel.chart.ScatterStyle.LINE_MARKER);

            var s1 = (XDDFScatterChartData.Series) scatterData.addSeries(xData, y1Data);
            s1.setTitle("Quadratic", null);
            s1.setMarkerStyle(org.apache.poi.xddf.usermodel.chart.MarkerStyle.CIRCLE);
            s1.setMarkerSize((short) 5);

            var s2 = (XDDFScatterChartData.Series) scatterData.addSeries(xData, y2Data);
            s2.setTitle("Linear", null);
            s2.setMarkerStyle(org.apache.poi.xddf.usermodel.chart.MarkerStyle.DIAMOND);
            s2.setMarkerSize((short) 5);

            chart.plot(scatterData);

            // Line styles via CT
            var ctScatter = chart.getCTChart().getPlotArea().getScatterChartArray(0);
            var ln1 = ctScatter.getSerArray(0).addNewSpPr().addNewLn();
            ln1.setW(19050); // 1.5pt
            ln1.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            var ln2 = ctScatter.getSerArray(1).addNewSpPr().addNewLn();
            ln2.setW(19050);
            ln2.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "scatter");
    }

    // ===========================
    // 6. Area Chart (Stacked)
    // ===========================
    @Test
    void compareAreaStacked() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_area.xlsx");
        Path poiFile = tempDir.resolve("poi_area.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.area()
                .title("Area Chart - Stacked")
                .position("E1:L15")
                .grouping(Grouping.STACKED)
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.solid("4472C4"))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .fill(ChartFill.solid("ED7D31"))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Area Chart - Stacked");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var costData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));

            XDDFAreaChartData areaData = (XDDFAreaChartData) chart.createData(ChartTypes.AREA, catAxis, valAxis);

            areaData.addSeries(cats, revData).setTitle("Revenue", null);
            areaData.addSeries(cats, costData).setTitle("Costs", null);

            chart.plot(areaData);

            // Set stacked grouping via CT
            chart.getCTChart().getPlotArea().getAreaChartArray(0).addNewGrouping()
                .setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STGrouping.STACKED);

            // Fills via CT
            var ctArea = chart.getCTChart().getPlotArea().getAreaChartArray(0);
            ctArea.getSerArray(0).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            ctArea.getSerArray(1).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "area_stacked");
    }

    // ===========================
    // 7. Doughnut Chart
    // ===========================
    @Test
    void compareDoughnut() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_doughnut.xlsx");
        Path poiFile = tempDir.resolve("poi_doughnut.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Market");
            writePieData(sheet);

            Chart chart = Chart.builder(ChartType.DOUGHNUT)
                .title("Doughnut Chart")
                .position("D1:K15")
                .addSeries(ChartSeries.builder()
                    .name("Share")
                    .categories("Market!$A$2:$A$6")
                    .values("Market!$B$2:$B$6")
                    .dataLabel(ChartDataLabel.builder()
                        .showValue(true)
                        .showCategory(true)
                        .separator(" - ")
                        .build())
                    .build())
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Market");
            writePoiPieData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 0, 10, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Doughnut Chart");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 5, 0, 0));
            var vals = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 5, 1, 1));

            XDDFChartData doughnutData = chart.createData(ChartTypes.DOUGHNUT, catAxis, valAxis);
            doughnutData.addSeries(cats, vals).setTitle("Share", null);

            chart.plot(doughnutData);

            // Data labels via CT
            var ctDoughnut = chart.getCTChart().getPlotArea().getDoughnutChartArray(0);
            var dLbls = ctDoughnut.getSerArray(0).addNewDLbls();
            dLbls.addNewShowVal().setVal(true);
            dLbls.addNewShowCatName().setVal(true);
            dLbls.addNewShowSerName().setVal(false);
            dLbls.addNewShowPercent().setVal(false);
            dLbls.addNewShowLegendKey().setVal(false);
            dLbls.setSeparator(" - ");

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "doughnut");
    }

    // ===========================
    // 8. Radar Chart (Filled)
    // ===========================
    @Test
    void compareRadar() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_radar.xlsx");
        Path poiFile = tempDir.resolve("poi_radar.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.builder(ChartType.RADAR)
                .title("Radar Chart")
                .position("E1:L15")
                .radarStyle(RadarStyle.FILLED)
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.solid("4472C4"))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .fill(ChartFill.solid("ED7D31"))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Radar Chart");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var costData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));

            XDDFRadarChartData radarData = (XDDFRadarChartData) chart.createData(ChartTypes.RADAR, catAxis, valAxis);
            radarData.setStyle(org.apache.poi.xddf.usermodel.chart.RadarStyle.FILLED);

            radarData.addSeries(cats, revData).setTitle("Revenue", null);
            radarData.addSeries(cats, costData).setTitle("Costs", null);

            chart.plot(radarData);

            // Fills via CT
            var ctRadar = chart.getCTChart().getPlotArea().getRadarChartArray(0);
            ctRadar.getSerArray(0).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            ctRadar.getSerArray(1).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "radar");
    }

    // ===========================
    // 9. 3D Column Chart
    // ===========================
    @Test
    void compare3DColumn() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_3d_column.xlsx");
        Path poiFile = tempDir.resolve("poi_3d_column.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.builder(ChartType.COLUMN_3D)
                .title("3D Column Chart")
                .position("E1:L15")
                .view3D(ChartView3D.of(20, 30, 40))
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.solid("4472C4"))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .fill(ChartFill.solid("ED7D31"))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("3D Column Chart");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var costData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));

            XDDFBar3DChartData bar3DData = (XDDFBar3DChartData) chart.createData(ChartTypes.BAR3D, catAxis, valAxis);
            bar3DData.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.COL);

            bar3DData.addSeries(cats, revData).setTitle("Revenue", null);
            bar3DData.addSeries(cats, costData).setTitle("Costs", null);

            chart.plot(bar3DData);

            // 3D view via CT
            var view3D = chart.getCTChart().addNewView3D();
            view3D.addNewRotX().setVal((byte) 20);
            view3D.addNewRotY().setVal(30);
            view3D.addNewPerspective().setVal((short) 40);

            // Fills via CT
            var ctBar3D = chart.getCTChart().getPlotArea().getBar3DChartArray(0);
            ctBar3D.getSerArray(0).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            ctBar3D.getSerArray(1).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "3d_column");
    }

    // ===========================
    // 10. Gradient Fill Chart
    // ===========================
    @Test
    void compareGradientFill() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_gradient.xlsx");
        Path poiFile = tempDir.resolve("poi_gradient.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.column()
                .title("Gradient Fill Demo")
                .position("E1:L15")
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.gradient(
                        ChartColor.rgb("4472C4"),
                        ChartColor.rgb("1F3864"),
                        90.0))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .fill(ChartFill.gradient(
                        List.of(
                            new GradientStop(0.0, ChartColor.rgb("ED7D31")),
                            new GradientStop(0.5, ChartColor.rgb("F4B183")),
                            new GradientStop(1.0, ChartColor.rgb("FBE5D6"))
                        ), 270.0))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Gradient Fill Demo");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var costData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));

            XDDFBarChartData barData = (XDDFBarChartData) chart.createData(ChartTypes.BAR, catAxis, valAxis);
            barData.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.COL);

            barData.addSeries(cats, revData).setTitle("Revenue", null);
            barData.addSeries(cats, costData).setTitle("Costs", null);

            chart.plot(barData);

            // Gradient fills via CT
            var ctBar = chart.getCTChart().getPlotArea().getBarChartArray(0);

            // Series 0: 2-stop gradient, 90 degrees
            var gradFill1 = ctBar.getSerArray(0).addNewSpPr().addNewGradFill();
            var gsLst1 = gradFill1.addNewGsLst();
            var gs1a = gsLst1.addNewGs();
            gs1a.setPos(0);
            gs1a.addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            var gs1b = gsLst1.addNewGs();
            gs1b.setPos(100000);
            gs1b.addNewSrgbClr().setVal(new byte[]{0x1F, 0x38, 0x64});
            var lin1 = gradFill1.addNewLin();
            lin1.setAng(5400000); // 90 * 60000
            lin1.setScaled(true);

            // Series 1: 3-stop gradient, 270 degrees
            var gradFill2 = ctBar.getSerArray(1).addNewSpPr().addNewGradFill();
            var gsLst2 = gradFill2.addNewGsLst();
            var gs2a = gsLst2.addNewGs();
            gs2a.setPos(0);
            gs2a.addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});
            var gs2b = gsLst2.addNewGs();
            gs2b.setPos(50000);
            gs2b.addNewSrgbClr().setVal(new byte[]{(byte) 0xF4, (byte) 0xB1, (byte) 0x83});
            var gs2c = gsLst2.addNewGs();
            gs2c.setPos(100000);
            gs2c.addNewSrgbClr().setVal(new byte[]{(byte) 0xFB, (byte) 0xE5, (byte) 0xD6});
            var lin2 = gradFill2.addNewLin();
            lin2.setAng(16200000); // 270 * 60000
            lin2.setScaled(true);

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "gradient");
    }

    // ===========================
    // 11. Pattern Fill Chart
    // ===========================
    @Test
    void comparePatternFill() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_pattern.xlsx");
        Path poiFile = tempDir.resolve("poi_pattern.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.column()
                .title("Pattern Fill Demo")
                .position("E1:L15")
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.pattern(PatternType.DIAGONAL_CROSS,
                        ChartColor.rgb("4472C4"), ChartColor.rgb("FFFFFF")))
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .fill(ChartFill.pattern(PatternType.HORIZONTAL,
                        ChartColor.rgb("ED7D31"), ChartColor.rgb("FFFFFF")))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Pattern Fill Demo");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var costData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));

            XDDFBarChartData barData = (XDDFBarChartData) chart.createData(ChartTypes.BAR, catAxis, valAxis);
            barData.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.COL);

            barData.addSeries(cats, revData).setTitle("Revenue", null);
            barData.addSeries(cats, costData).setTitle("Costs", null);

            chart.plot(barData);

            // Pattern fills via CT
            var ctBar = chart.getCTChart().getPlotArea().getBarChartArray(0);

            var pattFill1 = ctBar.getSerArray(0).addNewSpPr().addNewPattFill();
            pattFill1.setPrst(org.openxmlformats.schemas.drawingml.x2006.main.STPresetPatternVal.DIAG_CROSS);
            pattFill1.addNewFgClr().addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            pattFill1.addNewBgClr().addNewSrgbClr().setVal(new byte[]{(byte) 0xFF, (byte) 0xFF, (byte) 0xFF});

            var pattFill2 = ctBar.getSerArray(1).addNewSpPr().addNewPattFill();
            pattFill2.setPrst(org.openxmlformats.schemas.drawingml.x2006.main.STPresetPatternVal.HORZ);
            pattFill2.addNewFgClr().addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});
            pattFill2.addNewBgClr().addNewSrgbClr().setVal(new byte[]{(byte) 0xFF, (byte) 0xFF, (byte) 0xFF});

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "pattern");
    }

    // ===========================
    // 12. Custom Axes Configuration
    // ===========================
    @Test
    void compareCustomAxes() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_custom_axes.xlsx");
        Path poiFile = tempDir.resolve("poi_custom_axes.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.column()
                .title("Custom Axes Demo")
                .position("E1:L15")
                .categoryAxis(CategoryAxis.builder(0, 1)
                    .title("Month")
                    .position(AxisPosition.BOTTOM)
                    .majorTickMark(AxisTickMark.OUTSIDE)
                    .minorTickMark(AxisTickMark.INSIDE)
                    .build())
                .valueAxis(ValueAxis.builder(1, 0)
                    .title("Amount ($)")
                    .position(AxisPosition.LEFT)
                    .minimum(0)
                    .maximum(25000)
                    .majorUnit(5000)
                    .numberFormat("#,##0")
                    .crosses(AxisCrosses.AUTO_ZERO)
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.solid("4472C4"))
                    .dataLabel(ChartDataLabel.showValues())
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Custom Axes Demo");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            catAxis.setTitle("Month");
            catAxis.setMajorTickMark(org.apache.poi.xddf.usermodel.chart.AxisTickMark.OUT);
            catAxis.setMinorTickMark(org.apache.poi.xddf.usermodel.chart.AxisTickMark.IN);

            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setTitle("Amount ($)");
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);
            valAxis.setMinimum(0);
            valAxis.setMaximum(25000);
            valAxis.setMajorUnit(5000);
            valAxis.setNumberFormat("#,##0");

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));

            XDDFBarChartData barData = (XDDFBarChartData) chart.createData(ChartTypes.BAR, catAxis, valAxis);
            barData.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.COL);

            var s1 = barData.addSeries(cats, revData);
            s1.setTitle("Revenue", null);

            chart.plot(barData);

            // Solid fill via CT
            var ctBar = chart.getCTChart().getPlotArea().getBarChartArray(0);
            ctBar.getSerArray(0).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});

            // Data labels via CT
            var dLbls = ctBar.getSerArray(0).addNewDLbls();
            dLbls.addNewShowVal().setVal(true);
            dLbls.addNewShowCatName().setVal(false);
            dLbls.addNewShowSerName().setVal(false);
            dLbls.addNewShowPercent().setVal(false);
            dLbls.addNewShowLegendKey().setVal(false);

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "custom_axes");
    }

    // ===========================
    // 13. Line Chart with Error Bars
    // ===========================
    @Test
    void compareErrorBars() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_error_bars.xlsx");
        Path poiFile = tempDir.resolve("poi_error_bars.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.line()
                .title("Line with Error Bars")
                .position("E1:L15")
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .line(ChartLine.of("4472C4", 2.0))
                    .marker(MarkerStyle.CIRCLE, 5)
                    .errorBars(ChartErrorBars.percentage(10))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Line with Error Bars");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));

            XDDFLineChartData lineData = (XDDFLineChartData) chart.createData(ChartTypes.LINE, catAxis, valAxis);
            var s1 = (XDDFLineChartData.Series) lineData.addSeries(cats, revData);
            s1.setTitle("Revenue", null);
            s1.setMarkerStyle(org.apache.poi.xddf.usermodel.chart.MarkerStyle.CIRCLE);
            s1.setMarkerSize((short) 5);

            chart.plot(lineData);

            // Line style via CT
            var ctLine = chart.getCTChart().getPlotArea().getLineChartArray(0);
            var ln = ctLine.getSerArray(0).addNewSpPr().addNewLn();
            ln.setW(25400); // 2.0pt
            ln.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});

            // Error bars via CT
            var errBars = ctLine.getSerArray(0).addNewErrBars();
            errBars.addNewErrDir().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STErrDir.Y);
            errBars.addNewErrBarType().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STErrBarType.BOTH);
            errBars.addNewErrValType().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STErrValType.PERCENTAGE);
            errBars.addNewVal().setVal(10.0);

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "error_bars");
    }

    // ===========================
    // 14. Custom Title Font
    // ===========================
    @Test
    void compareCustomTitleFont() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_title_font.xlsx");
        Path poiFile = tempDir.resolve("poi_title_font.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            ChartFont titleFont = ChartFont.builder()
                .name("Arial")
                .size(16)
                .bold(true)
                .italic(false)
                .color(ChartColor.rgb("FF0000"))
                .build();

            Chart chart = Chart.column()
                .title(ChartTitle.of("Custom Font Title", titleFont))
                .position("E1:L15")
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.solid("4472C4"))
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Custom Font Title");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));

            XDDFBarChartData barData = (XDDFBarChartData) chart.createData(ChartTypes.BAR, catAxis, valAxis);
            barData.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.COL);

            barData.addSeries(cats, revData).setTitle("Revenue", null);

            chart.plot(barData);

            // Fill via CT
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0)
                .addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});

            // Custom font on title via CT
            var ctTitle = chart.getCTChart().getTitle();
            if (ctTitle != null && ctTitle.getTx() != null && ctTitle.getTx().getRich() != null) {
                var richText = ctTitle.getTx().getRich();
                if (richText.sizeOfPArray() > 0) {
                    var paragraph = richText.getPArray(0);
                    if (paragraph.sizeOfRArray() > 0) {
                        var run = paragraph.getRArray(0);
                        var rPr = run.isSetRPr() ? run.getRPr() : run.addNewRPr();
                        rPr.setSz(1600); // 16pt * 100
                        rPr.setB(true);
                        rPr.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 0xFF, 0x00, 0x00});
                        rPr.addNewLatin().setTypeface("Arial");
                    }
                }
            }

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "title_font");
    }

    // ===========================
    // 15. Data Labels with Values
    // ===========================
    @Test
    void compareDataLabels() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_data_labels.xlsx");
        Path poiFile = tempDir.resolve("poi_data_labels.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.column()
                .title("Data Labels Demo")
                .position("E1:L15")
                .addSeries(ChartSeries.builder()
                    .name("Revenue")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .fill(ChartFill.solid("4472C4"))
                    .dataLabel(ChartDataLabel.builder()
                        .showValue(true)
                        .showSeriesName(true)
                        .separator(" : ")
                        .build())
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Costs")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .fill(ChartFill.solid("ED7D31"))
                    .dataLabel(ChartDataLabel.showValues())
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Data Labels Demo");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var revData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var costData = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));

            XDDFBarChartData barData = (XDDFBarChartData) chart.createData(ChartTypes.BAR, catAxis, valAxis);
            barData.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.COL);

            barData.addSeries(cats, revData).setTitle("Revenue", null);
            barData.addSeries(cats, costData).setTitle("Costs", null);

            chart.plot(barData);

            var ctBar = chart.getCTChart().getPlotArea().getBarChartArray(0);
            ctBar.getSerArray(0).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            ctBar.getSerArray(1).addNewSpPr().addNewSolidFill()
                .addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});

            // Data labels series 0
            var dLbls0 = ctBar.getSerArray(0).addNewDLbls();
            dLbls0.addNewShowVal().setVal(true);
            dLbls0.addNewShowSerName().setVal(true);
            dLbls0.addNewShowCatName().setVal(false);
            dLbls0.addNewShowPercent().setVal(false);
            dLbls0.addNewShowLegendKey().setVal(false);
            dLbls0.setSeparator(" : ");

            // Data labels series 1
            var dLbls1 = ctBar.getSerArray(1).addNewDLbls();
            dLbls1.addNewShowVal().setVal(true);
            dLbls1.addNewShowSerName().setVal(false);
            dLbls1.addNewShowCatName().setVal(false);
            dLbls1.addNewShowPercent().setVal(false);
            dLbls1.addNewShowLegendKey().setVal(false);

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "data_labels");
    }

    // ===========================
    // 16. Line Dash Styles
    // ===========================
    @Test
    void compareLineDashStyles() throws Exception {
        Path litexlFile = tempDir.resolve("litexl_dash_styles.xlsx");
        Path poiFile = tempDir.resolve("poi_dash_styles.xlsx");

        // --- LiteXL ---
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            writeLiteXLData(sheet);

            Chart chart = Chart.line()
                .title("Line Dash Styles")
                .position("E1:L15")
                .addSeries(ChartSeries.builder()
                    .name("Solid")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$B$2:$B$7")
                    .line(ChartLine.builder().color("4472C4").width(2.0).dash(LineDash.SOLID).build())
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Dash")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$C$2:$C$7")
                    .line(ChartLine.builder().color("ED7D31").width(2.0).dash(LineDash.DASH).build())
                    .build())
                .addSeries(ChartSeries.builder()
                    .name("Dot")
                    .categories("Data!$A$2:$A$7")
                    .values("Data!$D$2:$D$7")
                    .line(ChartLine.builder().color("A5A5A5").width(2.0).dash(LineDash.DOT).build())
                    .build())
                .legend(LegendPosition.BOTTOM)
                .build();

            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // --- POI ---
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            writePoiData(sheet);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Line Dash Styles");

            XDDFCategoryAxis catAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            var cats = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 6, 0, 0));
            var d1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 1, 1));
            var d2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 2, 2));
            var d3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 6, 3, 3));

            XDDFLineChartData lineData = (XDDFLineChartData) chart.createData(ChartTypes.LINE, catAxis, valAxis);
            lineData.addSeries(cats, d1).setTitle("Solid", null);
            lineData.addSeries(cats, d2).setTitle("Dash", null);
            lineData.addSeries(cats, d3).setTitle("Dot", null);

            chart.plot(lineData);

            var ctLine = chart.getCTChart().getPlotArea().getLineChartArray(0);
            // Series 0: Solid
            var ln0 = ctLine.getSerArray(0).addNewSpPr().addNewLn();
            ln0.setW(25400);
            ln0.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0x44, 0x72, (byte) 0xC4});
            ln0.addNewPrstDash().setVal(org.openxmlformats.schemas.drawingml.x2006.main.STPresetLineDashVal.SOLID);

            // Series 1: Dash
            var ln1 = ctLine.getSerArray(1).addNewSpPr().addNewLn();
            ln1.setW(25400);
            ln1.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 0xED, 0x7D, 0x31});
            ln1.addNewPrstDash().setVal(org.openxmlformats.schemas.drawingml.x2006.main.STPresetLineDashVal.DASH);

            // Series 2: Dot
            var ln2 = ctLine.getSerArray(2).addNewSpPr().addNewLn();
            ln2.setW(25400);
            ln2.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte) 0xA5, (byte) 0xA5, (byte) 0xA5});
            ln2.addNewPrstDash().setVal(org.openxmlformats.schemas.drawingml.x2006.main.STPresetLineDashVal.DOT);

            chart.getOrAddLegend().setPosition(org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM);

            writePoiFile(wb, poiFile);
        }

        compareFiles(litexlFile, poiFile, "dash_styles");
    }

    // ===========================
    // Helper methods
    // ===========================

    private void writeLiteXLData(Sheet sheet) {
        sheet.cell(0, 0).set("Month");
        sheet.cell(0, 1).set("Revenue");
        sheet.cell(0, 2).set("Costs");
        sheet.cell(0, 3).set("Profit");
        for (int i = 0; i < CATEGORIES.length; i++) {
            sheet.cell(i + 1, 0).set(CATEGORIES[i]);
            sheet.cell(i + 1, 1).set(REVENUE[i]);
            sheet.cell(i + 1, 2).set(COSTS[i]);
            sheet.cell(i + 1, 3).set(PROFIT[i]);
        }
    }

    private void writePieData(Sheet sheet) {
        sheet.cell(0, 0).set("Product");
        sheet.cell(0, 1).set("Share");
        for (int i = 0; i < PIE_LABELS.length; i++) {
            sheet.cell(i + 1, 0).set(PIE_LABELS[i]);
            sheet.cell(i + 1, 1).set(PIE_VALUES[i]);
        }
    }

    private void writePoiData(XSSFSheet sheet) {
        var header = sheet.createRow(0);
        header.createCell(0).setCellValue("Month");
        header.createCell(1).setCellValue("Revenue");
        header.createCell(2).setCellValue("Costs");
        header.createCell(3).setCellValue("Profit");
        for (int i = 0; i < CATEGORIES.length; i++) {
            var row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(CATEGORIES[i]);
            row.createCell(1).setCellValue(REVENUE[i]);
            row.createCell(2).setCellValue(COSTS[i]);
            row.createCell(3).setCellValue(PROFIT[i]);
        }
    }

    private void writePoiPieData(XSSFSheet sheet) {
        var header = sheet.createRow(0);
        header.createCell(0).setCellValue("Product");
        header.createCell(1).setCellValue("Share");
        for (int i = 0; i < PIE_LABELS.length; i++) {
            var row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(PIE_LABELS[i]);
            row.createCell(1).setCellValue(PIE_VALUES[i]);
        }
    }

    private void writePoiFile(XSSFWorkbook wb, Path file) throws Exception {
        try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
            wb.write(fos);
        }
    }

    /**
     * Compares chart XML between LiteXL and POI files, copies to /tmp, verifies POI readability.
     */
    private void compareFiles(Path litexlFile, Path poiFile, String label) throws Exception {
        // Create output directory
        Files.createDirectories(OUTPUT_DIR);

        // Copy to /tmp for manual inspection
        Files.copy(litexlFile, OUTPUT_DIR.resolve("litexl_" + label + ".xlsx"), StandardCopyOption.REPLACE_EXISTING);
        Files.copy(poiFile, OUTPUT_DIR.resolve("poi_" + label + ".xlsx"), StandardCopyOption.REPLACE_EXISTING);

        System.out.println("\n" + "=".repeat(80));
        System.out.println("COMPARISON: " + label);
        System.out.println("=".repeat(80));

        // Print ZIP entries
        System.out.println("\n--- ZIP Entries ---");
        System.out.println("LiteXL:");
        printZipEntries(litexlFile);
        System.out.println("POI:");
        printZipEntries(poiFile);

        // Compare key XML files
        String[] xmlFiles = {
            "xl/charts/chart1.xml",
            "xl/drawings/drawing1.xml",
            "xl/worksheets/_rels/sheet1.xml.rels",
            "xl/drawings/_rels/drawing1.xml.rels",
        };

        for (String name : xmlFiles) {
            String litexlXml = readZipEntry(litexlFile, name);
            String poiXml = readZipEntry(poiFile, name);

            if (!"(not found)".equals(litexlXml) || !"(not found)".equals(poiXml)) {
                System.out.println("\n>>> " + name + " <<<");
                System.out.println("--- LiteXL ---");
                System.out.println(litexlXml);
                System.out.println("--- POI ---");
                System.out.println(poiXml);
            }
        }

        // Verify POI can read the LiteXL file
        try (FileInputStream fis = new FileInputStream(litexlFile.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            assertNotNull(drawing, label + ": Drawing should exist in LiteXL file");
            assertFalse(drawing.getCharts().isEmpty(), label + ": Should have charts in LiteXL file");

            System.out.println("\nPOI successfully read LiteXL file: " + drawing.getCharts().size() + " chart(s)");
        }

        System.out.println("\nFiles saved to: " + OUTPUT_DIR.resolve("litexl_" + label + ".xlsx"));
        System.out.println("                " + OUTPUT_DIR.resolve("poi_" + label + ".xlsx"));
    }

    private void printZipEntries(Path file) throws Exception {
        try (ZipFile zf = new ZipFile(file.toFile())) {
            zf.stream()
                .sorted(Comparator.comparing(ZipEntry::getName))
                .forEach(e -> System.out.println("  " + e.getName() + " (" + e.getSize() + " bytes)"));
        }
    }

    private String readZipEntry(Path file, String entryName) throws Exception {
        try (ZipFile zf = new ZipFile(file.toFile())) {
            ZipEntry entry = zf.getEntry(entryName);
            if (entry == null) return "(not found)";
            try (InputStream is = zf.getInputStream(entry)) {
                return new String(is.readAllBytes(), StandardCharsets.UTF_8);
            }
        }
    }
}
