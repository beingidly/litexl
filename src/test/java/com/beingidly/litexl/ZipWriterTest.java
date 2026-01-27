package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.io.IOException;
import java.nio.file.Path;
import java.util.zip.ZipFile;
import static org.junit.jupiter.api.Assertions.*;

class ZipWriterTest {

    @TempDir
    Path tempDir;

    @Test
    void writeEntry() throws IOException {
        Path zipPath = tempDir.resolve("output.zip");
        try (ZipWriter writer = new ZipWriter(zipPath)) {
            try (var os = writer.newEntry("content.txt")) {
                os.write("Test content".getBytes());
            }
        }

        // Verify
        try (ZipFile zf = new ZipFile(zipPath.toFile())) {
            var entry = zf.getEntry("content.txt");
            assertNotNull(entry);
            try (var is = zf.getInputStream(entry)) {
                assertEquals("Test content", new String(is.readAllBytes()));
            }
        }
    }

    @Test
    void writeMultipleEntries() throws IOException {
        Path zipPath = tempDir.resolve("multi.zip");
        try (ZipWriter writer = new ZipWriter(zipPath)) {
            try (var os = writer.newEntry("file1.txt")) {
                os.write("Content 1".getBytes());
            }
            try (var os = writer.newEntry("file2.txt")) {
                os.write("Content 2".getBytes());
            }
        }

        // Verify
        try (ZipFile zf = new ZipFile(zipPath.toFile())) {
            assertNotNull(zf.getEntry("file1.txt"));
            assertNotNull(zf.getEntry("file2.txt"));
        }
    }
}
