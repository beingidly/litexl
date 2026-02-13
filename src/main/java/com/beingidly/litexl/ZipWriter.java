package com.beingidly.litexl;

import org.jspecify.annotations.Nullable;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * ZIP file writer.
 */
final class ZipWriter implements AutoCloseable {

    private static final int ENTRY_BUFFER_SIZE = 65536;

    private final ZipOutputStream zipOut;
    private @Nullable BufferedEntryOutputStream currentEntryStream;

    public ZipWriter(java.nio.file.Path path) throws IOException {
        this.zipOut = new ZipOutputStream(new BufferedOutputStream(java.nio.file.Files.newOutputStream(path), 65536));
    }

    /**
     * Creates a ZIP writer that writes to an OutputStream.
     *
     * @param outputStream the output stream to write to
     * @throws IOException if an I/O error occurs
     */
    public ZipWriter(OutputStream outputStream) throws IOException {
        this.zipOut = new ZipOutputStream(new BufferedOutputStream(outputStream, 65536));
    }

    /**
     * Creates a new entry and returns an output stream to write to it.
     *
     * <p>Entry data is written directly to the ZIP stream without a temp file.</p>
     */
    public OutputStream newEntry(String name) throws IOException {
        // Close previous entry if any
        if (currentEntryStream != null) {
            currentEntryStream.closeEntry();
            currentEntryStream = null;
        }

        zipOut.putNextEntry(new ZipEntry(name));
        currentEntryStream = new BufferedEntryOutputStream(new EntryOutputStream());
        return currentEntryStream;
    }

    @Override
    public void close() throws IOException {
        if (currentEntryStream != null) {
            currentEntryStream.closeEntry();
            currentEntryStream = null;
        }
        zipOut.close();
    }

    /**
     * OutputStream for a single ZIP entry.
     * Closing this stream closes only the current entry.
     */
    private final class BufferedEntryOutputStream extends java.io.BufferedOutputStream {
        private final EntryOutputStream raw;
        private boolean closed;

        BufferedEntryOutputStream(EntryOutputStream raw) {
            super(raw, ENTRY_BUFFER_SIZE);
            this.raw = raw;
        }

        @Override
        public void close() throws IOException {
            if (closed) {
                return;
            }
            closed = true;
            super.close();
            if (currentEntryStream == this) {
                currentEntryStream = null;
            }
        }

        private void closeEntry() throws IOException {
            if (closed) {
                return;
            }
            closed = true;
            flush();
            raw.closeEntry();
        }
    }

    private final class EntryOutputStream extends OutputStream {
        private boolean closed;

        @Override
        public void write(int b) throws IOException {
            ensureOpen();
            zipOut.write(b);
        }

        @Override
        public void write(byte[] b, int off, int len) throws IOException {
            ensureOpen();
            zipOut.write(b, off, len);
        }

        @Override
        public void flush() throws IOException {
            zipOut.flush();
        }

        @Override
        public void close() throws IOException {
            closeEntry();
        }

        private void closeEntry() throws IOException {
            if (!closed) {
                zipOut.closeEntry();
                closed = true;
            }
        }

        private void ensureOpen() throws IOException {
            if (closed) {
                throw new IOException("ZIP entry stream is closed");
            }
        }
    }
}
