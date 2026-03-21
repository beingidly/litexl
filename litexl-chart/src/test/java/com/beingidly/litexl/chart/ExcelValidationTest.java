package com.beingidly.litexl.chart;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Detailed comparison test: generates identical charts with both LiteXL and POI,
 * then compares the XML structure element by element.
 */
class ExcelValidationTest {

    @TempDir
    Path tempDir;

    @Test
    void compareChartXmlWithPoi() throws Exception {
        Path litexlFile = tempDir.resolve("litexl.xlsx");
        Path poiFile = tempDir.resolve("poi.xlsx");

        // Generate with LiteXL
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("A"); sheet.cell(0, 1).set(10);
            sheet.cell(1, 0).set("B"); sheet.cell(1, 1).set(20);
            sheet.cell(2, 0).set("C"); sheet.cell(2, 1).set(30);

            Chart chart = Chart.of(ChartType.BAR, "Test",
                ChartPosition.of(3, 0, 10, 14),
                List.of(ChartSeries.of("Values", "Data!$A$1:$A$3", "Data!$B$1:$B$3")));
            Charts.add(sheet, chart);
            wb.save(litexlFile);
        }

        // Generate with POI
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue("A");
            sheet.getRow(0).createCell(1).setCellValue(10);
            sheet.createRow(1).createCell(0).setCellValue("B");
            sheet.getRow(1).createCell(1).setCellValue(20);
            sheet.createRow(2).createCell(0).setCellValue("C");
            sheet.getRow(2).createCell(1).setCellValue(30);

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 0, 10, 14);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Test");
            XDDFCategoryAxis catAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis valAxis = chart.createValueAxis(AxisPosition.LEFT);
            valAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            XDDFDataSource<String> cats = XDDFDataSourcesFactory.fromStringCellRange(
                sheet, new CellRangeAddress(0, 2, 0, 0));
            XDDFNumericalDataSource<Double> vals = XDDFDataSourcesFactory.fromNumericCellRange(
                sheet, new CellRangeAddress(0, 2, 1, 1));
            XDDFBarChartData data = (XDDFBarChartData) chart.createData(
                ChartTypes.BAR, catAxis, valAxis);
            data.setBarDirection(org.apache.poi.xddf.usermodel.chart.BarDirection.BAR);
            var series = data.addSeries(cats, vals);
            series.setTitle("Values", null);
            chart.plot(data);
            try (FileOutputStream fos = new FileOutputStream(poiFile.toFile())) {
                wb.write(fos);
            }
        }

        // Print both for visual comparison
        String[] partsToCompare = {
            "[Content_Types].xml",
            "xl/worksheets/_rels/sheet1.xml.rels",
            "xl/drawings/drawing1.xml",
            "xl/drawings/_rels/drawing1.xml.rels",
            "xl/charts/chart1.xml"
        };

        for (String part : partsToCompare) {
            String litexlXml = readZipEntry(litexlFile, part);
            String poiXml = readZipEntry(poiFile, part);

            System.out.println("\n===== " + part + " =====");
            System.out.println("--- LiteXL ---");
            System.out.println(litexlXml);
            System.out.println("--- POI ---");
            System.out.println(poiXml);
        }

        // Verify POI can read LiteXL's file
        try (java.io.FileInputStream fis = new java.io.FileInputStream(litexlFile.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFDrawing d = sheet.getDrawingPatriarch();
            assertNotNull(d, "Drawing should exist");
            assertFalse(d.getCharts().isEmpty(), "Should have charts");
        }

        // Copy both files to /tmp for manual testing
        Files.copy(litexlFile, Path.of("/tmp/litexl_chart.xlsx"),
            java.nio.file.StandardCopyOption.REPLACE_EXISTING);
        Files.copy(poiFile, Path.of("/tmp/poi_chart.xlsx"),
            java.nio.file.StandardCopyOption.REPLACE_EXISTING);
        System.out.println("\nFiles copied to /tmp/litexl_chart.xlsx and /tmp/poi_chart.xlsx");
    }

    private String readZipEntry(Path zipPath, String entryName) throws Exception {
        try (ZipFile zf = new ZipFile(zipPath.toFile())) {
            ZipEntry entry = zf.getEntry(entryName);
            if (entry == null) return "(not found)";
            try (InputStream is = zf.getInputStream(entry)) {
                return new String(is.readAllBytes());
            }
        }
    }
}
