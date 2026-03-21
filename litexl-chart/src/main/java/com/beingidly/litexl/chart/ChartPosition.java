package com.beingidly.litexl.chart;

import com.beingidly.litexl.CellRange;

/**
 * Position of a chart on a worksheet using a two-cell anchor.
 *
 * <p>Offsets are in EMU (English Metric Units). 1 inch = 914400 EMU.
 */
public record ChartPosition(
    int fromCol, int fromRow, int fromColOff, int fromRowOff,
    int toCol, int toRow, int toColOff, int toRowOff
) {

    /**
     * Creates a position from column/row indices (no offsets).
     */
    public static ChartPosition of(int fromCol, int fromRow, int toCol, int toRow) {
        return new ChartPosition(fromCol, fromRow, 0, 0, toCol, toRow, 0, 0);
    }

    /**
     * Creates a position from a cell range.
     */
    public static ChartPosition of(CellRange range) {
        return of(range.startCol(), range.startRow(), range.endCol(), range.endRow());
    }

    /**
     * Creates a position from a range string like "E1:L15".
     */
    public static ChartPosition of(String rangeRef) {
        return of(CellRange.parse(rangeRef));
    }
}
