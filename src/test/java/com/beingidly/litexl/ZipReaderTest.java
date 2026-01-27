package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import static org.junit.jupiter.api.Assertions.*;

class ZipReaderTest {

    @TempDir
    Path tempDir;

    @Test
    void readEntry() throws IOException {
        Path zipPath = createTestZip();
        try (ZipReader reader = new ZipReader(zipPath)) {
            try (var is = reader.getEntry("test.txt")) {
                assertNotNull(is);
                String content = new String(is.readAllBytes());
                assertEquals("Hello, World!", content);
            }
        }
    }

    @Test
    void getEntry_notFound() throws IOException {
        Path zipPath = createTestZip();
        try (ZipReader reader = new ZipReader(zipPath)) {
            assertNull(reader.getEntry("nonexistent.txt"));
        }
    }

    @Test
    void hasEntry_exists() throws IOException {
        Path zipPath = createTestZip();
        try (ZipReader reader = new ZipReader(zipPath)) {
            assertTrue(reader.hasEntry("test.txt"));
        }
    }

    @Test
    void hasEntry_notExists() throws IOException {
        Path zipPath = createTestZip();
        try (ZipReader reader = new ZipReader(zipPath)) {
            assertFalse(reader.hasEntry("nonexistent.txt"));
        }
    }

    private Path createTestZip() throws IOException {
        Path zipPath = tempDir.resolve("test.zip");
        try (var zos = new ZipOutputStream(Files.newOutputStream(zipPath))) {
            ZipEntry entry = new ZipEntry("test.txt");
            zos.putNextEntry(entry);
            zos.write("Hello, World!".getBytes());
            zos.closeEntry();
        }
        return zipPath;
    }
}
