package com.beingidly.litexl.benchmark;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.TimeUnit;

/**
 * JMH benchmark comparing read performance between litexl and Apache POI.
 */
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@State(Scope.Benchmark)
@Fork(1)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
public class ReadBenchmark {

    @Param({"1000", "10000", "50000"})
    private int rows;

    private static final int COLS = 10;
    private Path testFile;

    @Setup(Level.Trial)
    public void setup() throws Exception {
        testFile = Files.createTempFile("benchmark-read-", ".xlsx");

        // Create test file using POI (neutral ground)
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Data");

            for (int r = 0; r < rows; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < COLS; c++) {
                    if (c % 3 == 0) {
                        row.createCell(c).setCellValue("Text " + r + "-" + c);
                    } else if (c % 3 == 1) {
                        row.createCell(c).setCellValue(r * COLS + c + 0.5);
                    } else {
                        row.createCell(c).setCellValue(r % 2 == 0);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(testFile.toFile())) {
                wb.write(fos);
            }
        }
    }

    @TearDown(Level.Trial)
    public void tearDown() throws Exception {
        Files.deleteIfExists(testFile);
    }

    @Benchmark
    public void litexlRead(Blackhole bh) throws Exception {
        try (Workbook wb = Workbook.open(testFile)) {
            Sheet sheet = wb.getSheet(0);

            for (var row : sheet.rows().values()) {
                for (var cell : row.cells().values()) {
                    bh.consume(cell.rawValue());
                }
            }
        }
    }

    @Benchmark
    public void poiRead(Blackhole bh) throws Exception {
        try (FileInputStream fis = new FileInputStream(testFile.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            XSSFSheet sheet = wb.getSheetAt(0);

            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                var row = sheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c < row.getLastCellNum(); c++) {
                        var cell = row.getCell(c);
                        if (cell != null) {
                            bh.consume(cell.toString());
                        }
                    }
                }
            }
        }
    }

    @Benchmark
    public void litexlReadTyped(Blackhole bh) throws Exception {
        try (Workbook wb = Workbook.open(testFile)) {
            Sheet sheet = wb.getSheet(0);

            for (var row : sheet.rows().values()) {
                for (var cell : row.cells().values()) {
                    bh.consume(cell.value());
                    bh.consume(cell.type());
                }
            }
        }
    }

    @Benchmark
    public void poiReadTyped(Blackhole bh) throws Exception {
        try (FileInputStream fis = new FileInputStream(testFile.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            XSSFSheet sheet = wb.getSheetAt(0);

            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                var row = sheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c < row.getLastCellNum(); c++) {
                        var cell = row.getCell(c);
                        if (cell != null) {
                            bh.consume(cell.getCellType());
                            switch (cell.getCellType()) {
                                case STRING -> bh.consume(cell.getStringCellValue());
                                case NUMERIC -> bh.consume(cell.getNumericCellValue());
                                case BOOLEAN -> bh.consume(cell.getBooleanCellValue());
                                default -> bh.consume(cell.toString());
                            }
                        }
                    }
                }
            }
        }
    }
}
