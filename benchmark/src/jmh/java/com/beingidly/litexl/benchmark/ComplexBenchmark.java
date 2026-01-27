package com.beingidly.litexl.benchmark;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Random;
import java.util.concurrent.TimeUnit;

/**
 * JMH benchmark comparing complex workbook operations between litexl and Apache POI.
 *
 * <p>Benchmarks realistic workbook scenarios with multiple sheets, styles, formulas,
 * and various data types using the ComplexDataGenerator.
 */
@BenchmarkMode({Mode.AverageTime, Mode.Throughput})
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@State(Scope.Benchmark)
@Fork(1)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
public class ComplexBenchmark {

    private static final int MODIFY_CELLS = 300;

    private ComplexDataGenerator generator;
    private Path litexlWriteFile;
    private Path poiWriteFile;
    private Path litexlReadFile;
    private Path poiReadFile;
    private Path litexlModifyFile;
    private Path poiModifyFile;

    @Setup(Level.Trial)
    public void setupTrial() throws Exception {
        generator = new ComplexDataGenerator();

        // Create temp files for write benchmarks
        litexlWriteFile = Files.createTempFile("complex-litexl-write-", ".xlsx");
        poiWriteFile = Files.createTempFile("complex-poi-write-", ".xlsx");

        // Create temp files for read benchmarks
        litexlReadFile = Files.createTempFile("complex-litexl-read-", ".xlsx");
        poiReadFile = Files.createTempFile("complex-poi-read-", ".xlsx");

        // Create temp files for modify benchmarks
        litexlModifyFile = Files.createTempFile("complex-litexl-modify-", ".xlsx");
        poiModifyFile = Files.createTempFile("complex-poi-modify-", ".xlsx");

        // Pre-generate files for read benchmarks
        generator.resetRandom();
        generator.generateLitexl(litexlReadFile);
        generator.resetRandom();
        generator.generatePoi(poiReadFile);
    }

    @TearDown(Level.Trial)
    public void tearDownTrial() throws Exception {
        Files.deleteIfExists(litexlWriteFile);
        Files.deleteIfExists(poiWriteFile);
        Files.deleteIfExists(litexlReadFile);
        Files.deleteIfExists(poiReadFile);
        Files.deleteIfExists(litexlModifyFile);
        Files.deleteIfExists(poiModifyFile);
    }

    @Setup(Level.Invocation)
    public void setupInvocation() {
        generator.resetRandom();
    }

    // ========== Write Benchmarks ==========

    @Benchmark
    public void litexlWrite(Blackhole bh) throws Exception {
        generator.generateLitexl(litexlWriteFile);
        bh.consume(litexlWriteFile);
    }

    @Benchmark
    public void poiWrite(Blackhole bh) throws Exception {
        generator.generatePoi(poiWriteFile);
        bh.consume(poiWriteFile);
    }

    // ========== Read Benchmarks ==========

    @Benchmark
    public void litexlRead(Blackhole bh) throws Exception {
        try (Workbook wb = Workbook.open(litexlReadFile)) {
            for (int sheetIdx = 0; sheetIdx < wb.sheetCount(); sheetIdx++) {
                Sheet sheet = wb.getSheet(sheetIdx);
                bh.consume(sheet.name());
                for (var row : sheet.rows().values()) {
                    for (var cell : row.cells().values()) {
                        bh.consume(cell.rawValue());
                    }
                }
            }
        }
    }

    @Benchmark
    public void poiRead(Blackhole bh) throws Exception {
        try (FileInputStream fis = new FileInputStream(poiReadFile.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            for (int sheetIdx = 0; sheetIdx < wb.getNumberOfSheets(); sheetIdx++) {
                XSSFSheet sheet = wb.getSheetAt(sheetIdx);
                bh.consume(sheet.getSheetName());
                for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                    XSSFRow row = sheet.getRow(r);
                    if (row != null) {
                        for (int c = 0; c < row.getLastCellNum(); c++) {
                            XSSFCell cell = row.getCell(c);
                            if (cell != null) {
                                bh.consume(cell.toString());
                            }
                        }
                    }
                }
            }
        }
    }

    // ========== Read-Modify-Write Benchmarks ==========

    @Benchmark
    public void litexlReadModifyWrite(Blackhole bh) throws Exception {
        // Copy source file for modification
        Files.copy(litexlReadFile, litexlModifyFile, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        try (Workbook wb = Workbook.open(litexlModifyFile)) {
            Random rand = new Random(42);
            Sheet sheet = wb.getSheet(1); // Sales sheet (largest)

            // Modify 300 random cells
            for (int i = 0; i < MODIFY_CELLS; i++) {
                int row = rand.nextInt(3000) + 1; // Skip header
                int col = rand.nextInt(5); // First 5 columns (non-formula)
                sheet.cell(row, col).set("Modified-" + i);
            }

            wb.save(litexlModifyFile);
        }
        bh.consume(litexlModifyFile);
    }

    @Benchmark
    public void poiReadModifyWrite(Blackhole bh) throws Exception {
        // Copy source file for modification
        Files.copy(poiReadFile, poiModifyFile, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        try (FileInputStream fis = new FileInputStream(poiModifyFile.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            Random rand = new Random(42);
            XSSFSheet sheet = wb.getSheetAt(1); // Sales sheet (largest)

            // Modify 300 random cells
            for (int i = 0; i < MODIFY_CELLS; i++) {
                int rowNum = rand.nextInt(3000) + 1; // Skip header
                int colNum = rand.nextInt(5); // First 5 columns (non-formula)
                XSSFRow row = sheet.getRow(rowNum);
                if (row != null) {
                    XSSFCell cell = row.getCell(colNum);
                    if (cell == null) {
                        cell = row.createCell(colNum);
                    }
                    cell.setCellValue("Modified-" + i);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(poiModifyFile.toFile())) {
                wb.write(fos);
            }
        }
        bh.consume(poiModifyFile);
    }
}
