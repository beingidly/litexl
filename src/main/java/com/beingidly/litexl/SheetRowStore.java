package com.beingidly.litexl;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.EOFException;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.LocalDateTime;

/**
 * File-backed row store for streaming sheets.
 */
final class SheetRowStore implements AutoCloseable {

    @FunctionalInterface
    interface RowVisitor {
        /**
         * @return true to continue iteration, false to stop
         */
        boolean visit(Row row);
    }

    private static final byte TYPE_EMPTY = 0;
    private static final byte TYPE_STRING = 1;
    private static final byte TYPE_NUMBER = 2;
    private static final byte TYPE_BOOLEAN = 3;
    private static final byte TYPE_DATE = 4;
    private static final byte TYPE_FORMULA = 5;
    private static final byte TYPE_ERROR = 6;

    private final Path path;
    private final DataOutputStream out;
    private boolean sealed;
    private boolean closed;

    SheetRowStore() {
        try {
            this.path = Files.createTempFile("litexl-sheet-", ".rows");
            this.out = new DataOutputStream(new BufferedOutputStream(
                Files.newOutputStream(path, StandardOpenOption.WRITE, StandardOpenOption.TRUNCATE_EXISTING)));
        } catch (IOException e) {
            throw new LitexlException(ErrorCode.IO_ERROR, "Failed to create sheet row store", e);
        }
    }

    void append(Row row) {
        ensureWritable();
        try {
            out.writeInt(row.rowNum());
            out.writeDouble(row.height());
            out.writeBoolean(row.hasCustomHeight());
            out.writeBoolean(row.hidden());

            out.writeInt(row.cellCount());
            for (Cell cell : row.cells().values()) {
                out.writeInt(cell.column());
                out.writeInt(cell.styleId());
                writeCellValue(cell.value());
            }
        } catch (IOException e) {
            throw new LitexlException(ErrorCode.IO_ERROR, "Failed to append row to sheet store", e);
        }
    }

    void seal() {
        if (closed || sealed) {
            return;
        }
        try {
            out.flush();
            out.close();
            sealed = true;
        } catch (IOException e) {
            throw new LitexlException(ErrorCode.IO_ERROR, "Failed to seal sheet row store", e);
        }
    }

    void forEachRow(RowVisitor visitor) {
        flushForRead();

        try (DataInputStream in = new DataInputStream(new BufferedInputStream(Files.newInputStream(path)))) {
            while (true) {
                Row row;
                try {
                    row = readRow(in);
                } catch (EOFException e) {
                    break;
                }
                if (!visitor.visit(row)) {
                    break;
                }
            }
        } catch (IOException e) {
            throw new LitexlException(ErrorCode.IO_ERROR, "Failed to read rows from sheet store", e);
        }
    }

    private Row readRow(DataInputStream in) throws IOException {
        int rowNum = in.readInt();
        double height = in.readDouble();
        boolean customHeight = in.readBoolean();
        boolean hidden = in.readBoolean();

        Row row = new Row(rowNum);
        if (customHeight) {
            row.height(height);
        }
        row.hidden(hidden);

        int cellCount = in.readInt();
        for (int i = 0; i < cellCount; i++) {
            int col = in.readInt();
            int styleId = in.readInt();

            Cell cell = row.cell(col);
            if (styleId > 0) {
                cell.style(styleId);
            }

            byte type = in.readByte();
            switch (type) {
                case TYPE_EMPTY -> cell.setEmpty();
                case TYPE_STRING -> cell.set(readString(in));
                case TYPE_NUMBER -> cell.set(in.readDouble());
                case TYPE_BOOLEAN -> cell.set(in.readBoolean());
                case TYPE_DATE -> cell.set(LocalDateTime.parse(readString(in)));
                case TYPE_FORMULA -> cell.setFormula(readString(in));
                case TYPE_ERROR -> cell.setValue(new CellValue.Error(readString(in)));
                default -> throw new IOException("Unknown cell type in row store: " + type);
            }
        }

        return row;
    }

    private void writeCellValue(CellValue value) throws IOException {
        switch (value) {
            case CellValue.Empty _ -> out.writeByte(TYPE_EMPTY);
            case CellValue.Text t -> {
                out.writeByte(TYPE_STRING);
                writeString(out, t.value());
            }
            case CellValue.Number n -> {
                out.writeByte(TYPE_NUMBER);
                out.writeDouble(n.value());
            }
            case CellValue.Bool b -> {
                out.writeByte(TYPE_BOOLEAN);
                out.writeBoolean(b.value());
            }
            case CellValue.Date d -> {
                out.writeByte(TYPE_DATE);
                writeString(out, d.value().toString());
            }
            case CellValue.Formula f -> {
                out.writeByte(TYPE_FORMULA);
                writeString(out, f.expression());
            }
            case CellValue.Error e -> {
                out.writeByte(TYPE_ERROR);
                writeString(out, e.code());
            }
        }
    }

    private void flushForRead() {
        if (closed) {
            throw new IllegalStateException("Sheet row store is closed");
        }
        if (sealed) {
            return;
        }
        try {
            out.flush();
        } catch (IOException e) {
            throw new LitexlException(ErrorCode.IO_ERROR, "Failed to flush sheet row store", e);
        }
    }

    private static void writeString(DataOutputStream out, String value) throws IOException {
        byte[] bytes = value.getBytes(StandardCharsets.UTF_8);
        out.writeInt(bytes.length);
        out.write(bytes);
    }

    private static String readString(DataInputStream in) throws IOException {
        int len = in.readInt();
        if (len < 0) {
            throw new IOException("Negative string length in row store");
        }
        byte[] bytes = in.readNBytes(len);
        if (bytes.length != len) {
            throw new EOFException("Unexpected EOF while reading row store string");
        }
        return new String(bytes, StandardCharsets.UTF_8);
    }

    private void ensureWritable() {
        if (closed) {
            throw new IllegalStateException("Sheet row store is closed");
        }
        if (sealed) {
            throw new IllegalStateException("Sheet row store is sealed");
        }
    }

    @Override
    public void close() {
        if (closed) {
            return;
        }

        try {
            if (!sealed) {
                out.close();
            }
        } catch (IOException e) {
            throw new LitexlException(ErrorCode.IO_ERROR, "Failed to close sheet row store", e);
        } finally {
            closed = true;
            try {
                Files.deleteIfExists(path);
            } catch (IOException e) {
                throw new LitexlException(ErrorCode.IO_ERROR, "Failed to delete sheet row store", e);
            }
        }
    }
}
