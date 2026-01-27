package com.beingidly.litexl;

import org.jspecify.annotations.Nullable;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.logging.Logger;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * ZIP file writer.
 */
final class ZipWriter implements AutoCloseable {

    private static final Logger logger = Logger.getLogger(ZipWriter.class.getName());

    private final ZipOutputStream zipOut;
    private final Path tempDir;
    private @Nullable TempFileOutputStream currentEntryStream;

    public ZipWriter(Path path) throws IOException {
        this.tempDir = Files.createTempDirectory("litexl-");
        this.zipOut = new ZipOutputStream(new BufferedOutputStream(Files.newOutputStream(path), 65536));
    }

    /**
     * Creates a ZIP writer that writes to an OutputStream.
     *
     * @param outputStream the output stream to write to
     * @throws IOException if an I/O error occurs
     */
    public ZipWriter(OutputStream outputStream) throws IOException {
        this.tempDir = Files.createTempDirectory("litexl-");
        this.zipOut = new ZipOutputStream(new BufferedOutputStream(outputStream, 65536));
    }

    /**
     * Creates a new entry and returns an output stream to write to it.
     * Writes to a temp file first, then copies to ZIP on close for optimal performance.
     */
    public OutputStream newEntry(String name) throws IOException {
        // Close and flush previous entry if any
        if (currentEntryStream != null) {
            currentEntryStream.flushToZip();
            currentEntryStream = null;
        }

        currentEntryStream = new TempFileOutputStream(name);
        return currentEntryStream;
    }

    @Override
    public void close() throws IOException {
        if (currentEntryStream != null) {
            currentEntryStream.flushToZip();
            currentEntryStream = null;
        }
        zipOut.close();

        // Clean up temp directory
        if (Files.exists(tempDir)) {
            try (Stream<Path> stream = Files.walk(tempDir)) {
                stream.sorted((a, b) -> -a.compareTo(b))
                    .forEach(p -> {
                        try {
                            Files.delete(p);
                        } catch (IOException e) {
                            logger.warning("Failed to delete temp file: " + p + " - " + e.getMessage());
                        }
                    });
            }
        }
    }

    /**
     * Writes to a temp file, then transfers to ZIP on close.
     * Avoids memory issues with large content and StAX/ZIP interaction overhead.
     */
    private class TempFileOutputStream extends BufferedOutputStream {
        private final String entryName;
        private final Path tempFile;
        private boolean flushed = false;

        TempFileOutputStream(String entryName) throws IOException {
            super(Files.newOutputStream(tempDir.resolve(
                entryName.replace('/', '_').replace('\\', '_'))), 65536);
            this.entryName = entryName;
            this.tempFile = tempDir.resolve(entryName.replace('/', '_').replace('\\', '_'));
        }

        void flushToZip() throws IOException {
            if (!flushed) {
                flushed = true;
                super.close(); // Close temp file first

                zipOut.putNextEntry(new ZipEntry(entryName));
                Files.copy(tempFile, zipOut);
                zipOut.closeEntry();

                Files.deleteIfExists(tempFile);
            }
        }

        @Override
        public void close() throws IOException {
            flushToZip();
        }
    }
}
