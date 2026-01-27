package com.beingidly.litexl;

import org.jspecify.annotations.Nullable;

import java.util.Collections;
import java.util.Map;
import java.util.NavigableMap;
import java.util.TreeMap;

/**
 * Represents a row in a worksheet.
 *
 * <p>This class is <b>not thread-safe</b>.
 * External synchronization is required for concurrent access.
 */
public final class Row {

    private final int rowNum;
    private final NavigableMap<Integer, Cell> cells;
    private double height;
    private boolean hidden;
    private boolean customHeight;

    Row(int rowNum) {
        this.rowNum = rowNum;
        this.cells = new TreeMap<>();
        this.height = -1; // Default height (auto)
        this.hidden = false;
        this.customHeight = false;
    }

    /**
     * Returns the row number (0-based).
     */
    public int rowNum() {
        return rowNum;
    }

    /**
     * Gets a cell by column index, creating it if it doesn't exist.
     *
     * @param col the 0-based column index
     * @throws IllegalArgumentException if column index is out of Excel's limits
     */
    public Cell cell(int col) {
        ExcelLimits.validateColumnIndex(col);
        return cells.computeIfAbsent(col, Cell::new);
    }

    /**
     * Gets a cell by column index, returning null if it doesn't exist.
     */
    public @Nullable Cell getCell(int col) {
        return cells.get(col);
    }

    /**
     * Returns all cells in this row (unmodifiable).
     */
    public Map<Integer, Cell> cells() {
        return Collections.unmodifiableMap(cells);
    }

    /**
     * Returns the number of cells in this row.
     */
    public int cellCount() {
        return cells.size();
    }

    /**
     * Returns the first column index, or {@link ColumnIndex#none()} if no cells.
     */
    public ColumnIndex firstColumn() {
        return cells.isEmpty() ? ColumnIndex.none() : ColumnIndex.of(cells.firstKey());
    }

    /**
     * Returns the last column index, or {@link ColumnIndex#none()} if no cells.
     */
    public ColumnIndex lastColumn() {
        return cells.isEmpty() ? ColumnIndex.none() : ColumnIndex.of(cells.lastKey());
    }

    /**
     * Sets the row height in points.
     */
    public Row height(double height) {
        this.height = height;
        this.customHeight = height >= 0;
        return this;
    }

    /**
     * Returns the row height in points, or -1 for auto height.
     */
    public double height() {
        return height;
    }

    /**
     * Returns true if this row has a custom height.
     */
    public boolean hasCustomHeight() {
        return customHeight;
    }

    /**
     * Sets whether this row is hidden.
     */
    public Row hidden(boolean hidden) {
        this.hidden = hidden;
        return this;
    }

    /**
     * Returns true if this row is hidden.
     */
    public boolean hidden() {
        return hidden;
    }

    @Override
    public String toString() {
        return String.format("Row[num=%d, cells=%d]", rowNum, cells.size());
    }
}
