package com.beingidly.litexl;

import com.beingidly.litexl.crypto.SheetProtection;
import com.beingidly.litexl.format.AutoFilter;
import com.beingidly.litexl.format.ConditionalFormat;
import com.beingidly.litexl.format.DataValidation;
import org.jspecify.annotations.Nullable;

import java.util.*;

/**
 * Represents a worksheet in a workbook.
 *
 * <p>This class is <b>not thread-safe</b>.
 * External synchronization is required for concurrent access.
 */
public final class Sheet {

    private String name;
    private final int index;
    private final NavigableMap<Integer, Row> rows;
    private final SheetFormat format;
    private final SheetProtectionManager protectionManager;

    /**
     * Creates a new sheet. Internal use only - use Workbook.addSheet() instead.
     * @hidden
     */
    public Sheet(String name, int index) {
        this.name = name;
        this.index = index;
        this.rows = new TreeMap<>();
        this.format = new SheetFormat();
        this.protectionManager = new SheetProtectionManager();
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
     * Gets a cell, creating row and cell if they don't exist.
     *
     * @param row the 0-based row index
     * @param col the 0-based column index
     * @throws IllegalArgumentException if indices are out of Excel's limits
     */
    public Cell cell(int row, int col) {
        ExcelLimits.validateCellIndex(row, col);
        return row(row).cell(col);
    }

    /**
     * Gets a cell, returning null if row or cell doesn't exist.
     */
    public @Nullable Cell getCell(int row, int col) {
        Row r = getRow(row);
        return r == null ? null : r.getCell(col);
    }

    /**
     * Gets a row, creating it if it doesn't exist.
     *
     * @param rowNum the 0-based row index
     * @throws IllegalArgumentException if row index is out of Excel's limits
     */
    public Row row(int rowNum) {
        ExcelLimits.validateRowIndex(rowNum);
        return rows.computeIfAbsent(rowNum, Row::new);
    }

    /**
     * Gets a row, returning null if it doesn't exist.
     */
    public @Nullable Row getRow(int rowNum) {
        return rows.get(rowNum);
    }

    /**
     * Returns all rows (unmodifiable).
     */
    public Map<Integer, Row> rows() {
        return Collections.unmodifiableMap(rows);
    }

    /**
     * Returns the number of rows.
     */
    public int rowCount() {
        return rows.size();
    }

    /**
     * Returns the first row index, or {@link RowIndex#none()} if no rows.
     */
    public RowIndex firstRow() {
        return rows.isEmpty() ? RowIndex.none() : RowIndex.of(rows.firstKey());
    }

    /**
     * Returns the last row index, or {@link RowIndex#none()} if no rows.
     */
    public RowIndex lastRow() {
        return rows.isEmpty() ? RowIndex.none() : RowIndex.of(rows.lastKey());
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
    public List<MergedRegion> mergedCells() {
        // Convert SheetFormat.MergedRegion to Sheet.MergedRegion for backward compatibility
        List<MergedRegion> result = new ArrayList<>();
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
    public List<ConditionalFormat> conditionalFormats() {
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
    public List<DataValidation> validations() {
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

    @Override
    public String toString() {
        return String.format("Sheet[name=%s, rows=%d]", name, rows.size());
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
