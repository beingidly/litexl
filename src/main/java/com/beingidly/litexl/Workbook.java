package com.beingidly.litexl;

import com.beingidly.litexl.crypto.EncryptionOptions;
import com.beingidly.litexl.style.Style;

import org.jspecify.annotations.Nullable;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

/**
 * Represents an Excel workbook.
 *
 * <p>A workbook contains one or more sheets and can be saved to or loaded
 * from an XLSX file.</p>
 *
 * <p>Workbook implements {@link AutoCloseable} and should be used with
 * try-with-resources to ensure proper resource cleanup:</p>
 *
 * <pre>{@code
 * try (Workbook wb = Workbook.create()) {
 *     Sheet sheet = wb.addSheet("Data");
 *     sheet.cell(0, 0).set("Hello");
 *     wb.save(Path.of("output.xlsx"));
 * }
 * }</pre>
 *
 * <p>This class is <b>not thread-safe</b>.
 * External synchronization is required for concurrent access.
 */
public final class Workbook implements AutoCloseable {

    private final List<Sheet> sheets;
    private final List<Style> styles;
    private final List<String> sharedStrings;
    private final Map<String, Integer> sharedStringIndex;
    private boolean closed;

    private Workbook() {
        this.sheets = new ArrayList<>();
        this.styles = new ArrayList<>();
        this.sharedStrings = new ArrayList<>();
        this.sharedStringIndex = new HashMap<>();
        this.closed = false;

        // Add default style at index 0
        styles.add(Style.DEFAULT);
    }

    /**
     * Creates a new empty workbook.
     */
    public static Workbook create() {
        return new Workbook();
    }

    /**
     * Opens an existing workbook from a file.
     */
    public static Workbook open(Path path) {
        return open(path, null);
    }

    /**
     * Opens an existing workbook from a file with a password.
     */
    public static Workbook open(Path path, @Nullable String password) {
        if (!Files.exists(path)) {
            throw new LitexlException(ErrorCode.FILE_NOT_FOUND, "File not found: " + path);
        }
        try (XlsxReader reader = new XlsxReader(path)) {
            return reader.read(password);
        } catch (IOException e) {
            throw new LitexlException(ErrorCode.IO_ERROR, "Failed to read file: " + path, e);
        }
    }

    /**
     * Saves the workbook to a file.
     */
    public void save(Path path) {
        save(path, null);
    }

    /**
     * Saves the workbook to a file with encryption.
     */
    public void save(Path path, @Nullable EncryptionOptions options) {
        ensureOpen();
        try (XlsxWriter writer = new XlsxWriter(this, path)) {
            if (options != null) {
                writer.withEncryption(options);
            }
            writer.write();
        } catch (IOException e) {
            throw new LitexlException(ErrorCode.IO_ERROR, "Failed to save file: " + path, e);
        }
    }

    /**
     * Saves the workbook to an output stream.
     *
     * <p>This is useful for streaming directly to HTTP responses in web applications:</p>
     * <pre>{@code
     * response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
     * try (Workbook wb = Workbook.create()) {
     *     Sheet sheet = wb.addSheet("Data");
     *     sheet.cell(0, 0).set("Hello");
     *     wb.save(response.getOutputStream());
     * }
     * }</pre>
     *
     * @param outputStream the output stream to write to
     */
    public void save(OutputStream outputStream) {
        save(outputStream, null);
    }

    /**
     * Saves the workbook to an output stream with encryption.
     *
     * @param outputStream the output stream to write to
     * @param options the encryption options, or null for no encryption
     */
    public void save(OutputStream outputStream, @Nullable EncryptionOptions options) {
        ensureOpen();
        try (XlsxWriter writer = new XlsxWriter(this, outputStream)) {
            if (options != null) {
                writer.withEncryption(options);
            }
            writer.write();
        } catch (IOException e) {
            throw new LitexlException(ErrorCode.IO_ERROR, "Failed to save to stream", e);
        }
    }

    /**
     * Adds a new sheet to the workbook.
     */
    public Sheet addSheet(String name) {
        ensureOpen();
        if (name.isEmpty()) {
            throw new IllegalArgumentException("Sheet name cannot be empty");
        }

        // Check for duplicate name
        for (Sheet sheet : sheets) {
            if (sheet.name().equalsIgnoreCase(name)) {
                throw new IllegalArgumentException("Sheet with name '" + name + "' already exists");
            }
        }

        Sheet sheet = new Sheet(name, sheets.size());
        sheets.add(sheet);
        return sheet;
    }

    /**
     * Gets a sheet by index, or null if index is out of range.
     */
    public @Nullable Sheet getSheet(int index) {
        ensureOpen();
        if (index < 0 || index >= sheets.size()) {
            return null;
        }
        return sheets.get(index);
    }

    /**
     * Gets a sheet by name, or null if not found.
     */
    public @Nullable Sheet getSheet(String name) {
        ensureOpen();
        for (Sheet sheet : sheets) {
            if (sheet.name().equalsIgnoreCase(name)) {
                return sheet;
            }
        }
        return null;
    }

    /**
     * Removes a sheet by index.
     */
    public void removeSheet(int index) {
        ensureOpen();
        if (index < 0 || index >= sheets.size()) {
            throw new IndexOutOfBoundsException("Sheet index out of range: " + index);
        }
        sheets.remove(index);
    }

    /**
     * Returns all sheets (unmodifiable).
     */
    public List<Sheet> sheets() {
        return Collections.unmodifiableList(sheets);
    }

    /**
     * Returns the number of sheets.
     */
    public int sheetCount() {
        return sheets.size();
    }

    /**
     * Adds a style and returns its ID.
     */
    public int addStyle(Style style) {
        ensureOpen();
        styles.add(style);
        return styles.size() - 1;
    }

    /**
     * Gets a style by ID, or null if not found.
     */
    public @Nullable Style getStyle(int id) {
        if (id < 0 || id >= styles.size()) {
            return null;
        }
        return styles.get(id);
    }

    /**
     * Returns all styles (unmodifiable).
     */
    public List<Style> styles() {
        return Collections.unmodifiableList(styles);
    }

    // === Shared Strings (internal use) ===

    /**
     * Adds a shared string and returns its index.
     */
    public int addSharedString(String value) {
        Integer existing = sharedStringIndex.get(value);
        if (existing != null) {
            return existing;
        }
        int index = sharedStrings.size();
        sharedStrings.add(value);
        sharedStringIndex.put(value, index);
        return index;
    }

    /**
     * Gets a shared string by index, or null if not found.
     */
    public @Nullable String getSharedString(int index) {
        if (index < 0 || index >= sharedStrings.size()) {
            return null;
        }
        return sharedStrings.get(index);
    }

    /**
     * Returns all shared strings (unmodifiable).
     */
    public List<String> sharedStrings() {
        return Collections.unmodifiableList(sharedStrings);
    }

    @Override
    public void close() {
        closed = true;
    }

    private void ensureOpen() {
        if (closed) {
            throw new IllegalStateException("Workbook is closed");
        }
    }

    @Override
    public String toString() {
        return String.format("Workbook[sheets=%d]", sheets.size());
    }

    /**
     * Adds a sheet directly for internal use by XlsxReader.
     */
    void addSheet(Sheet sheet) {
        sheets.add(sheet);
    }
}
