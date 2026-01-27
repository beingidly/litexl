package com.beingidly.litexl;

import org.jspecify.annotations.Nullable;

import java.io.*;
import java.nio.file.Path;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * ZIP file reader.
 */
final class ZipReader implements AutoCloseable {

    private final ZipFile zipFile;

    public ZipReader(Path path) throws IOException {
        this.zipFile = new ZipFile(path.toFile());
    }

    /**
     * Gets an entry's input stream, or null if not found.
     */
    public @Nullable InputStream getEntry(String name) throws IOException {
        ZipEntry entry = zipFile.getEntry(name);
        if (entry == null) {
            return null;
        }
        return zipFile.getInputStream(entry);
    }

    /**
     * Checks if an entry exists.
     */
    public boolean hasEntry(String name) {
        return zipFile.getEntry(name) != null;
    }

    @Override
    public void close() throws IOException {
        zipFile.close();
    }
}
