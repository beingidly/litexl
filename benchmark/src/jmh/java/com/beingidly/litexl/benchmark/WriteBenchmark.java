package com.beingidly.litexl.benchmark;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.TimeUnit;

/**
 * JMH benchmark comparing write performance between litexl and Apache POI.
 */
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@State(Scope.Benchmark)
@Fork(1)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
public class WriteBenchmark {

    @Param({"1000", "10000", "50000"})
    private int rows;

    private static final int COLS = 10;
    private Path litexlFile;
    private Path poiFile;

    @Setup(Level.Trial)
    public void setup() throws Exception {
        litexlFile = Files.createTempFile("benchmark-litexl-", ".xlsx");
        poiFile = Files.createTempFile("benchmark-poi-", ".xlsx");
    }

    @TearDown(Level.Trial)
    public void tearDown() throws Exception {
        Files.deleteIfExists(litexlFile);
        Files.deleteIfExists(poiFile);
    }

    @Benchmark
    public void litexlWrite() throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");

            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < COLS; c++) {
                    sheet.cell(r, c).set("Cell " + r + "-" + c);
                }
            }

            wb.save(litexlFile);
        }
    }

    @Benchmark
    public void poiWrite() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");

            for (int r = 0; r < rows; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < COLS; c++) {
                    row.createCell(c).setCellValue("Cell " + r + "-" + c);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(poiFile.toFile())) {
                wb.write(fos);
            }
        }
    }

    @Benchmark
    public void litexlWriteNumeric() throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Numbers");

            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < COLS; c++) {
                    sheet.cell(r, c).set(r * COLS + c + 0.5);
                }
            }

            wb.save(litexlFile);
        }
    }

    @Benchmark
    public void poiWriteNumeric() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Numbers");

            for (int r = 0; r < rows; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < COLS; c++) {
                    row.createCell(c).setCellValue(r * COLS + c + 0.5);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(poiFile.toFile())) {
                wb.write(fos);
            }
        }
    }

    @Benchmark
    public void litexlWriteMixed() throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Mixed");

            for (int r = 0; r < rows; r++) {
                sheet.cell(r, 0).set("Row " + r);
                sheet.cell(r, 1).set(r);
                sheet.cell(r, 2).set(r * 1.5);
                sheet.cell(r, 3).set(r % 2 == 0);
                sheet.cell(r, 4).setFormula("B" + (r + 1) + "+C" + (r + 1));
            }

            wb.save(litexlFile);
        }
    }

    @Benchmark
    public void poiWriteMixed() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Mixed");

            for (int r = 0; r < rows; r++) {
                var row = sheet.createRow(r);
                row.createCell(0).setCellValue("Row " + r);
                row.createCell(1).setCellValue(r);
                row.createCell(2).setCellValue(r * 1.5);
                row.createCell(3).setCellValue(r % 2 == 0);
                row.createCell(4).setCellFormula("B" + (r + 1) + "+C" + (r + 1));
            }

            try (FileOutputStream fos = new FileOutputStream(poiFile.toFile())) {
                wb.write(fos);
            }
        }
    }
}
