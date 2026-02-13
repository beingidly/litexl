package com.beingidly.litexl;

import com.beingidly.litexl.crypto.SheetProtection;
import com.beingidly.litexl.format.AutoFilter;
import com.beingidly.litexl.format.ConditionalFormat;
import com.beingidly.litexl.format.DataValidation;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Map;
import java.util.TreeMap;

/**
 * Represents a worksheet in a workbook.
 *
 * <p>This class is <b>not thread-safe</b>.
 * External synchronization is required for concurrent access.
 */
public final class Sheet {

    @FunctionalInterface
    public interface RowVisitor {
        /**
         * @param row current row
         * @return true to continue iteration, false to stop
         */
        boolean visit(Row row);
    }

    private String name;
    private final int index;
    private final SheetRowStore rowStore;
    private final SheetFormat format;
    private final SheetProtectionManager protectionManager;

    private @Nullable Row currentRow;
    private int rowCount;
    private int firstRow = -1;
    private int lastRow = -1;
    private boolean readOnly;

    /**
     * Creates a new sheet. Internal use only - use Workbook.addSheet() instead.
     * @hidden
     */
    public Sheet(String name, int index) {
        this.name = name;
        this.index = index;
        this.rowStore = new SheetRowStore();
        this.format = new SheetFormat();
        this.protectionManager = new SheetProtectionManager();
        this.currentRow = null;
        this.rowCount = 0;
        this.readOnly = false;
    }

    /**
     * Returns the sheet name.
     */
    public String name() {
        return name;
    }

    /**
     * Sets the sheet name.
     */
    public void setName(String name) {
        if (name.isEmpty()) {
            throw new IllegalArgumentException("Sheet name cannot be empty");
        }
        this.name = name;
    }

    /**
     * Returns the sheet index (0-based).
     */
    public int index() {
        return index;
    }

    /**
     * Gets a cell in append-only streaming mode.
     *
     * <p>Rows must be written in strictly increasing order. Once a row is flushed,
     * it cannot be modified again.</p>
     *
     * @param row the 0-based row index
     * @param col the 0-based column index
     * @throws IllegalArgumentException if indices are out of Excel's limits
     */
    public Cell cell(int row, int col) {
        ExcelLimits.validateCellIndex(row, col);

        if (readOnly) {
            Cell existing = getCell(row, col);
            if (existing == null) {
                throw new UnsupportedOperationException(
                    "Sheet is read-only in streaming mode. Cell does not exist: (" + row + "," + col + ")");
            }
            return existing;
        }

        if (canAppendToRow(row)) {
            return row(row).cell(col);
        }

        Cell existing = getCell(row, col);
        if (existing != null) {
            return existing;
        }
        throw new IllegalStateException(
            "Cannot create cells in flushed rows in streaming mode. Requested row=" + row + ", lastRow=" + lastRow);
    }

    /**
     * Gets a cell, returning null if row or cell doesn't exist.
     */
    public @Nullable Cell getCell(int row, int col) {
        Row r = getRow(row);
        return r == null ? null : r.getCell(col);
    }

    /**
     * Gets a row in append-only streaming mode.
     *
     * @param rowNum the 0-based row index
     * @throws IllegalArgumentException if row index is out of Excel's limits
     */
    public Row row(int rowNum) {
        ExcelLimits.validateRowIndex(rowNum);

        if (readOnly) {
            Row existing = getRow(rowNum);
            if (existing == null) {
                throw new UnsupportedOperationException(
                    "Sheet is read-only in streaming mode. Row does not exist: " + rowNum);
            }
            return existing;
        }

        if (currentRow != null) {
            int current = currentRow.rowNum();
            if (rowNum < current) {
                Row existing = getRow(rowNum);
                if (existing != null) {
                    return existing;
                }
                throw new IllegalStateException(
                    "Streaming sheets are append-only by row index. Current row=" + current + ", requested row=" + rowNum);
            }
            if (rowNum > current) {
                flushCurrentRow();
                currentRow = newRow(rowNum);
            }
            return currentRow;
        }

        if (lastRow >= 0 && rowNum <= lastRow) {
            Row existing = getRow(rowNum);
            if (existing != null) {
                return existing;
            }
            throw new IllegalStateException(
                "Streaming sheets are append-only by row index. Last flushed row=" + lastRow + ", requested row=" + rowNum);
        }

        currentRow = newRow(rowNum);
        return currentRow;
    }

    /**
     * Gets a row, returning null if it doesn't exist.
     */
    public @Nullable Row getRow(int rowNum) {
        if (rowNum < 0) {
            return null;
        }

        if (currentRow != null && currentRow.rowNum() == rowNum) {
            return currentRow;
        }

        final Row[] found = new Row[1];
        rowStore.forEachRow(row -> {
            if (row.rowNum() == rowNum) {
                found[0] = row;
                return false;
            }
            return row.rowNum() < rowNum;
        });
        return found[0];
    }

    /**
     * Iterates rows in ascending row index order without keeping all rows in memory.
     */
    public void forEachRow(RowVisitor visitor) {
        final boolean[] stopped = { false };
        rowStore.forEachRow(row -> {
            boolean keepGoing = visitor.visit(row);
            if (!keepGoing) {
                stopped[0] = true;
            }
            return keepGoing;
        });

        if (!stopped[0] && currentRow != null) {
            visitor.visit(currentRow);
        }
    }

    /**
     * Materializes all rows as a map.
     *
     * <p>This allocates memory proportional to sheet size and is not recommended
     * for large sheets. Prefer {@link #forEachRow(RowVisitor)} for streaming access.</p>
     */
    public Map<Integer, Row> rows() {
        Map<Integer, Row> result = new TreeMap<>();
        forEachRow(row -> {
            result.put(row.rowNum(), row);
            return true;
        });
        return Collections.unmodifiableMap(result);
    }

    /**
     * Returns the number of rows.
     */
    public int rowCount() {
        return rowCount;
    }

    /**
     * Returns the first row index, or {@link RowIndex#none()} if no rows.
     */
    public RowIndex firstRow() {
        return rowCount == 0 ? RowIndex.none() : RowIndex.of(firstRow);
    }

    /**
     * Returns the last row index, or {@link RowIndex#none()} if no rows.
     */
    public RowIndex lastRow() {
        return rowCount == 0 ? RowIndex.none() : RowIndex.of(lastRow);
    }

    // === Format Delegation ===

    /**
     * Returns the format manager for this sheet.
     */
    public SheetFormat format() {
        return format;
    }

    // === Protection Delegation ===

    /**
     * Returns the protection manager for this sheet.
     */
    public SheetProtectionManager protectionManager() {
        return protectionManager;
    }

    // === Convenience Methods (delegating to SheetFormat) ===

    /**
     * Merges a range of cells.
     */
    public void merge(int r1, int c1, int r2, int c2) {
        format.merge(r1, c1, r2, c2);
    }

    /**
     * Merges a range of cells.
     */
    public void merge(CellRange range) {
        format.merge(range);
    }

    /**
     * Unmerges a range of cells.
     */
    public void unmerge(int r1, int c1, int r2, int c2) {
        format.unmerge(r1, c1, r2, c2);
    }

    /**
     * Returns all merged regions (unmodifiable).
     */
    public java.util.List<MergedRegion> mergedCells() {
        // Convert SheetFormat.MergedRegion to Sheet.MergedRegion for backward compatibility
        java.util.List<MergedRegion> result = new ArrayList<>();
        for (SheetFormat.MergedRegion m : format.mergedCells()) {
            result.add(new MergedRegion(m.startRow(), m.startCol(), m.endRow(), m.endCol()));
        }
        return Collections.unmodifiableList(result);
    }

    /**
     * Sets an auto filter on the given range.
     */
    public void setAutoFilter(int r1, int c1, int r2, int c2) {
        format.setAutoFilter(r1, c1, r2, c2);
    }

    /**
     * Sets an auto filter on the given range.
     */
    public void setAutoFilter(CellRange range) {
        format.setAutoFilter(range);
    }

    /**
     * Sets an auto filter with specified filter columns.
     */
    public void setAutoFilter(AutoFilter filter) {
        format.setAutoFilter(filter);
    }

    /**
     * Clears the auto filter.
     */
    public void clearAutoFilter() {
        format.clearAutoFilter();
    }

    /**
     * Returns the auto filter, or null if not set.
     */
    public @Nullable AutoFilter autoFilter() {
        return format.autoFilter();
    }

    /**
     * Adds a conditional format.
     */
    public void addConditionalFormat(ConditionalFormat cf) {
        format.addConditionalFormat(cf);
    }

    /**
     * Returns all conditional formats (unmodifiable).
     */
    public java.util.List<ConditionalFormat> conditionalFormats() {
        return format.conditionalFormats();
    }

    /**
     * Clears all conditional formats.
     */
    public void clearConditionalFormats() {
        format.clearConditionalFormats();
    }

    /**
     * Adds a data validation.
     */
    public void addValidation(DataValidation dv) {
        format.addValidation(dv);
    }

    /**
     * Returns all data validations (unmodifiable).
     */
    public java.util.List<DataValidation> validations() {
        return format.validations();
    }

    /**
     * Clears all data validations.
     */
    public void clearValidations() {
        format.clearValidations();
    }

    /**
     * Sets the width of a column in characters.
     *
     * @param col the 0-based column index
     * @param width the width in characters
     * @throws IllegalArgumentException if column index is out of Excel's limits
     */
    public void setColumnWidth(int col, double width) {
        format.setColumnWidth(col, width);
    }

    /**
     * Returns the width of a column, or {@link ColumnWidth#auto()} if not set.
     */
    public ColumnWidth getColumnWidth(int col) {
        return format.getColumnWidth(col);
    }

    /**
     * Returns the raw column width, or -1 for auto width.
     * For internal use during save.
     */
    double getColumnWidthRaw(int col) {
        return format.getColumnWidthRaw(col);
    }

    /**
     * Returns all column widths (unmodifiable).
     */
    public Map<Integer, Double> columnWidths() {
        return format.columnWidths();
    }

    /**
     * Sets whether this sheet is hidden.
     */
    public void setHidden(boolean hidden) {
        format.setHidden(hidden);
    }

    /**
     * Returns true if this sheet is hidden.
     */
    public boolean hidden() {
        return format.hidden();
    }

    // === Convenience Methods (delegating to SheetProtectionManager) ===

    /**
     * Protects the sheet without a password.
     */
    public void protect(SheetProtection options) {
        protectionManager.protect(options);
    }

    /**
     * Returns the protection settings, or null if not protected.
     */
    public @Nullable SheetProtection protection() {
        return protectionManager.options();
    }

    /**
     * Returns true if the sheet is protected.
     */
    public boolean isProtected() {
        return protectionManager.isProtected();
    }

    void finishLoadingReadOnly() {
        flushCurrentRow();
        rowStore.seal();
        readOnly = true;
    }

    void closeResources() {
        rowStore.close();
    }

    @Override
    public String toString() {
        return String.format("Sheet[name=%s, rows=%d]", name, rowCount);
    }

    private Row newRow(int rowNum) {
        Row row = new Row(rowNum);
        rowCount++;
        if (firstRow < 0) {
            firstRow = rowNum;
        }
        lastRow = rowNum;
        return row;
    }

    private void flushCurrentRow() {
        if (currentRow == null) {
            return;
        }
        rowStore.append(currentRow);
        currentRow = null;
    }

    private boolean canAppendToRow(int rowNum) {
        if (currentRow != null) {
            return rowNum >= currentRow.rowNum();
        }
        return lastRow < 0 || rowNum > lastRow;
    }

    /**
     * Represents a merged cell region.
     */
    public record MergedRegion(int startRow, int startCol, int endRow, int endCol) {
        public CellRange toRange() {
            return CellRange.of(startRow, startCol, endRow, endCol);
        }
    }
}
