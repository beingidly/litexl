package com.beingidly.litexl.benchmark;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.crypto.EncryptionOptions;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.TimeUnit;

/**
 * JMH benchmark for measuring encryption performance.
 * Used to establish baseline before ByteBuffer migration.
 */
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@State(Scope.Benchmark)
@Fork(value = 1, jvmArgs = {"-Xms256m", "-Xmx1g", "-XX:+UseG1GC"})
@Warmup(iterations = 1, time = 2)
@Measurement(iterations = 2, time = 3)
public class EncryptBenchmark {

    @Param({"1000", "10000"})
    private int rows;

    private static final int COLS = 10;
    private static final String PASSWORD = "benchmark123";

    private Path encryptedFile;
    private EncryptionOptions encryptionOptions;

    @Setup(Level.Trial)
    public void setup() throws Exception {
        encryptedFile = Files.createTempFile("benchmark-encrypted-", ".xlsx");
        // Use low spinCount for benchmark speed (production uses 100000)
        encryptionOptions = new EncryptionOptions(
            EncryptionOptions.Algorithm.AES_256,
            PASSWORD,
            1000
        );
    }

    @TearDown(Level.Trial)
    public void tearDown() throws Exception {
        Files.deleteIfExists(encryptedFile);
    }

    @Benchmark
    public void writeEncrypted(Blackhole bh) throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("EncryptedData");

            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < COLS; c++) {
                    sheet.cell(r, c).set("Cell " + r + "-" + c);
                }
            }

            wb.save(encryptedFile, encryptionOptions);
        }

        bh.consume(Files.size(encryptedFile));
    }

    @Benchmark
    public void writeEncryptedNumeric(Blackhole bh) throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("EncryptedNumbers");

            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < COLS; c++) {
                    sheet.cell(r, c).set(r * COLS + c + 0.5);
                }
            }

            wb.save(encryptedFile, encryptionOptions);
        }

        bh.consume(Files.size(encryptedFile));
    }
}
