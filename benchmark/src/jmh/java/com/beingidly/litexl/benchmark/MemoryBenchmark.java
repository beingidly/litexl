package com.beingidly.litexl.benchmark;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.TimeUnit;

/**
 * JMH benchmark comparing memory efficiency between litexl and Apache POI.
 * Uses gc profiler to measure allocations.
 */
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@State(Scope.Benchmark)
@Fork(value = 1, jvmArgs = {"-Xms512m", "-Xmx512m"})
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
public class MemoryBenchmark {

    @Param({"10000", "50000"})
    private int rows;

    private static final int COLS = 20;
    private Path litexlFile;
    private Path poiFile;

    @Setup(Level.Trial)
    public void setup() throws Exception {
        litexlFile = Files.createTempFile("benchmark-mem-litexl-", ".xlsx");
        poiFile = Files.createTempFile("benchmark-mem-poi-", ".xlsx");
    }

    @TearDown(Level.Trial)
    public void tearDown() throws Exception {
        Files.deleteIfExists(litexlFile);
        Files.deleteIfExists(poiFile);
    }

    @Benchmark
    public void litexlLargeWrite(Blackhole bh) throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Large");

            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < COLS; c++) {
                    sheet.cell(r, c).set("Data-" + r + "-" + c);
                }
            }

            wb.save(litexlFile);
            bh.consume(wb);
        }
    }

    @Benchmark
    public void poiLargeWrite(Blackhole bh) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Large");

            for (int r = 0; r < rows; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < COLS; c++) {
                    row.createCell(c).setCellValue("Data-" + r + "-" + c);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(poiFile.toFile())) {
                wb.write(fos);
            }
            bh.consume(wb);
        }
    }

    @Benchmark
    public void litexlMultiSheet(Blackhole bh) throws Exception {
        try (Workbook wb = Workbook.create()) {
            for (int s = 0; s < 10; s++) {
                Sheet sheet = wb.addSheet("Sheet" + s);

                for (int r = 0; r < rows / 10; r++) {
                    for (int c = 0; c < COLS; c++) {
                        sheet.cell(r, c).set("S" + s + "-R" + r + "-C" + c);
                    }
                }
            }

            wb.save(litexlFile);
            bh.consume(wb);
        }
    }

    @Benchmark
    public void poiMultiSheet(Blackhole bh) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            for (int s = 0; s < 10; s++) {
                XSSFSheet sheet = wb.createSheet("Sheet" + s);

                for (int r = 0; r < rows / 10; r++) {
                    var row = sheet.createRow(r);
                    for (int c = 0; c < COLS; c++) {
                        row.createCell(c).setCellValue("S" + s + "-R" + r + "-C" + c);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(poiFile.toFile())) {
                wb.write(fos);
            }
            bh.consume(wb);
        }
    }

    @Benchmark
    public void litexlNumericData(Blackhole bh) throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Numeric");

            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < COLS; c++) {
                    sheet.cell(r, c).set((double) (r * COLS + c));
                }
            }

            wb.save(litexlFile);
            bh.consume(wb);
        }
    }

    @Benchmark
    public void poiNumericData(Blackhole bh) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Numeric");

            for (int r = 0; r < rows; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < COLS; c++) {
                    row.createCell(c).setCellValue((double) (r * COLS + c));
                }
            }

            try (FileOutputStream fos = new FileOutputStream(poiFile.toFile())) {
                wb.write(fos);
            }
            bh.consume(wb);
        }
    }
}
