package com.beingidly.litexl.chart;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * Generates a reference chart using POI and dumps the XML for comparison.
 */
class PoiReferenceChartGenerator {

    @TempDir
    Path tempDir;

    @Test
    void dumpPoiChartXml() throws Exception {
        Path file = tempDir.resolve("poi_chart.xlsx");

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

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                wb.write(fos);
            }
        }

        // Dump key XML entries
        try (ZipFile zf = new ZipFile(file.toFile())) {
            for (String name : new String[]{
                "xl/charts/chart1.xml",
                "xl/drawings/drawing1.xml"
            }) {
                ZipEntry entry = zf.getEntry(name);
                if (entry != null) {
                    System.out.println("\n===== " + name + " =====");
                    try (InputStream is = zf.getInputStream(entry)) {
                        System.out.println(new String(is.readAllBytes()));
                    }
                }
            }
        }
    }
}
