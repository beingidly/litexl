package com.beingidly.litexl;

import com.beingidly.litexl.format.AutoFilter;
import com.beingidly.litexl.format.ConditionalFormat;
import com.beingidly.litexl.format.DataValidation;
import org.jspecify.annotations.Nullable;

import java.util.*;

/**
 * Manages formatting settings for a worksheet.
 *
 * <p>This class handles merge cells, auto filters, conditional formatting,
 * data validation, and column widths.
 *
 * <p>This class is <b>not thread-safe</b>.
 * External synchronization is required for concurrent access.
 */
public final class SheetFormat {

    private final List<MergedRegion> mergedCells;
    private final List<ConditionalFormat> conditionalFormats;
    private final List<DataValidation> validations;
    private final NavigableMap<Integer, Double> columnWidths;
    private AutoFilter autoFilter;
    private boolean hidden;

    SheetFormat() {
        this.mergedCells = new ArrayList<>();
        this.conditionalFormats = new ArrayList<>();
        this.validations = new ArrayList<>();
        this.columnWidths = new TreeMap<>();
        this.hidden = false;
    }

    // === Merge Cells ===

    /**
     * Merges a range of cells.
     */
    public void merge(int r1, int c1, int r2, int c2) {
        mergedCells.add(new MergedRegion(r1, c1, r2, c2));
    }

    /**
     * Merges a range of cells.
     */
    public void merge(CellRange range) {
        mergedCells.add(new MergedRegion(range.startRow(), range.startCol(), range.endRow(), range.endCol()));
    }

    /**
     * Unmerges a range of cells.
     */
    public void unmerge(int r1, int c1, int r2, int c2) {
        mergedCells.removeIf(m ->
            m.startRow() == r1 && m.startCol() == c1 &&
            m.endRow() == r2 && m.endCol() == c2
        );
    }

    /**
     * Returns all merged regions (unmodifiable).
     */
    public List<MergedRegion> mergedCells() {
        return Collections.unmodifiableList(mergedCells);
    }

    // === AutoFilter ===

    /**
     * Sets an auto filter on the given range.
     */
    public void setAutoFilter(int r1, int c1, int r2, int c2) {
        this.autoFilter = new AutoFilter(CellRange.of(r1, c1, r2, c2), List.of());
    }

    /**
     * Sets an auto filter on the given range.
     */
    public void setAutoFilter(CellRange range) {
        this.autoFilter = new AutoFilter(range, List.of());
    }

    /**
     * Sets an auto filter with specified filter columns.
     */
    public void setAutoFilter(AutoFilter filter) {
        this.autoFilter = filter;
    }

    /**
     * Clears the auto filter.
     */
    public void clearAutoFilter() {
        this.autoFilter = null;
    }

    /**
     * Returns the auto filter, or null if not set.
     */
    public @Nullable AutoFilter autoFilter() {
        return autoFilter;
    }

    // === Conditional Formatting ===

    /**
     * Adds a conditional format.
     */
    public void addConditionalFormat(ConditionalFormat cf) {
        conditionalFormats.add(cf);
    }

    /**
     * Returns all conditional formats (unmodifiable).
     */
    public List<ConditionalFormat> conditionalFormats() {
        return Collections.unmodifiableList(conditionalFormats);
    }

    /**
     * Clears all conditional formats.
     */
    public void clearConditionalFormats() {
        conditionalFormats.clear();
    }

    // === Data Validation ===

    /**
     * Adds a data validation.
     */
    public void addValidation(DataValidation dv) {
        validations.add(dv);
    }

    /**
     * Returns all data validations (unmodifiable).
     */
    public List<DataValidation> validations() {
        return Collections.unmodifiableList(validations);
    }

    /**
     * Clears all data validations.
     */
    public void clearValidations() {
        validations.clear();
    }

    // === Column Width ===

    /**
     * Sets the width of a column in characters.
     *
     * @param col the 0-based column index
     * @param width the width in characters
     * @throws IllegalArgumentException if column index is out of Excel's limits
     */
    public void setColumnWidth(int col, double width) {
        ExcelLimits.validateColumnIndex(col);
        columnWidths.put(col, width);
    }

    /**
     * Returns the width of a column, or {@link ColumnWidth#auto()} if not set.
     */
    public ColumnWidth getColumnWidth(int col) {
        Double width = columnWidths.get(col);
        return width != null ? ColumnWidth.of(width) : ColumnWidth.auto();
    }

    /**
     * Returns the raw column width, or -1 for auto width.
     * For internal use during save.
     */
    double getColumnWidthRaw(int col) {
        return columnWidths.getOrDefault(col, -1.0);
    }

    /**
     * Returns all column widths (unmodifiable).
     */
    public Map<Integer, Double> columnWidths() {
        return Collections.unmodifiableMap(columnWidths);
    }

    // === Hidden ===

    /**
     * Sets whether this sheet is hidden.
     */
    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    /**
     * Returns true if this sheet is hidden.
     */
    public boolean hidden() {
        return hidden;
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
