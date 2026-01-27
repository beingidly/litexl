package com.beingidly.litexl.benchmark;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.style.*;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.*;
import org.openjdk.jmh.annotations.*;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.TimeUnit;

/**
 * JMH benchmark comparing styled cell operations between litexl and Apache POI.
 */
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@State(Scope.Benchmark)
@Fork(1)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
public class StyleBenchmark {

    @Param({"1000", "5000"})
    private int rows;

    private static final int COLS = 10;
    private Path litexlFile;
    private Path poiFile;

    @Setup(Level.Trial)
    public void setup() throws Exception {
        litexlFile = Files.createTempFile("benchmark-litexl-style-", ".xlsx");
        poiFile = Files.createTempFile("benchmark-poi-style-", ".xlsx");
    }

    @TearDown(Level.Trial)
    public void tearDown() throws Exception {
        Files.deleteIfExists(litexlFile);
        Files.deleteIfExists(poiFile);
    }

    @Benchmark
    public void litexlWriteStyled() throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Styled");

            // Register styles
            int boldStyle = wb.addStyle(Style.builder().bold(true).build());
            int fillStyle = wb.addStyle(Style.builder().fill(0xFFFFFF00).build());
            int borderStyle = wb.addStyle(Style.builder()
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < COLS; c++) {
                    var cell = sheet.cell(r, c).set("Cell " + r + "-" + c);

                    if (r == 0) {
                        cell.style(boldStyle);
                    } else if (c == 0) {
                        cell.style(fillStyle);
                    } else {
                        cell.style(borderStyle);
                    }
                }
            }

            wb.save(litexlFile);
        }
    }

    @Benchmark
    public void poiWriteStyled() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Styled");

            // Create styles
            XSSFCellStyle boldStyle = wb.createCellStyle();
            XSSFFont boldFont = wb.createFont();
            boldFont.setBold(true);
            boldStyle.setFont(boldFont);

            XSSFCellStyle fillStyle = wb.createCellStyle();
            fillStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 255, (byte) 255, 0}, null));
            fillStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            XSSFCellStyle borderStyle = wb.createCellStyle();
            borderStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            borderStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            borderStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            borderStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);

            for (int r = 0; r < rows; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < COLS; c++) {
                    var cell = row.createCell(c);
                    cell.setCellValue("Cell " + r + "-" + c);

                    if (r == 0) {
                        cell.setCellStyle(boldStyle);
                    } else if (c == 0) {
                        cell.setCellStyle(fillStyle);
                    } else {
                        cell.setCellStyle(borderStyle);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(poiFile.toFile())) {
                wb.write(fos);
            }
        }
    }

    @Benchmark
    public void litexlWriteComplexStyles() throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Complex");

            // Create multiple complex styles
            int headerStyle = wb.addStyle(Style.builder()
                .bold(true)
                .fill(0xFF4472C4)
                .color(0xFFFFFFFF)
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .border(BorderStyle.MEDIUM, 0xFF000000)
                .build());

            int numberStyle = wb.addStyle(Style.builder()
                .format("#,##0.00")
                .align(HAlign.RIGHT, VAlign.BOTTOM)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            int dateStyle = wb.addStyle(Style.builder()
                .format("yyyy-mm-dd")
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .build());

            // Header row
            for (int c = 0; c < COLS; c++) {
                sheet.cell(0, c).set("Header " + c).style(headerStyle);
            }

            // Data rows
            for (int r = 1; r < rows; r++) {
                for (int c = 0; c < COLS; c++) {
                    var cell = sheet.cell(r, c);
                    if (c % 2 == 0) {
                        cell.set(r * c * 1.5).style(numberStyle);
                    } else {
                        cell.set(45000 + r).style(dateStyle);
                    }
                }
            }

            wb.save(litexlFile);
        }
    }

    @Benchmark
    public void poiWriteComplexStyles() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Complex");

            // Header style
            XSSFCellStyle headerStyle = wb.createCellStyle();
            XSSFFont headerFont = wb.createFont();
            headerFont.setBold(true);
            headerFont.setColor(new XSSFColor(new byte[]{(byte) 255, (byte) 255, (byte) 255}, null));
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 68, (byte) 114, (byte) 196}, null));
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM);
            headerStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM);
            headerStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM);
            headerStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.MEDIUM);

            // Number style
            XSSFCellStyle numberStyle = wb.createCellStyle();
            numberStyle.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
            numberStyle.setAlignment(HorizontalAlignment.RIGHT);
            numberStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
            numberStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            numberStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            numberStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            numberStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);

            // Date style
            XSSFCellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd"));
            dateStyle.setAlignment(HorizontalAlignment.CENTER);
            dateStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // Header row
            var headerRow = sheet.createRow(0);
            for (int c = 0; c < COLS; c++) {
                var cell = headerRow.createCell(c);
                cell.setCellValue("Header " + c);
                cell.setCellStyle(headerStyle);
            }

            // Data rows
            for (int r = 1; r < rows; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < COLS; c++) {
                    var cell = row.createCell(c);
                    if (c % 2 == 0) {
                        cell.setCellValue(r * c * 1.5);
                        cell.setCellStyle(numberStyle);
                    } else {
                        cell.setCellValue(45000 + r);
                        cell.setCellStyle(dateStyle);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(poiFile.toFile())) {
                wb.write(fos);
            }
        }
    }
}
