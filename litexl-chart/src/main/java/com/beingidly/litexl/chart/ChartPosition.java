package com.beingidly.litexl.chart;

import com.beingidly.litexl.CellRange;

/**
 * Position of a chart on a worksheet using a two-cell anchor.
 *
 * <p>Offsets are in EMU (English Metric Units). 1 inch = 914400 EMU.
 *
 * @param fromCol starting column index (zero-based)
 * @param fromRow starting row index (zero-based)
 * @param fromColOff column offset in EMU from the starting column
 * @param fromRowOff row offset in EMU from the starting row
 * @param toCol ending column index (zero-based)
 * @param toRow ending row index (zero-based)
 * @param toColOff column offset in EMU from the ending column
 * @param toRowOff row offset in EMU from the ending row
 */
public record ChartPosition(
    int fromCol, int fromRow, int fromColOff, int fromRowOff,
    int toCol, int toRow, int toColOff, int toRowOff
) {

    /**
     * Creates a position from column/row indices (no offsets).
     *
     * @param fromCol starting column index
     * @param fromRow starting row index
     * @param toCol ending column index
     * @param toRow ending row index
     * @return a new chart position
     */
    public static ChartPosition of(int fromCol, int fromRow, int toCol, int toRow) {
        return new ChartPosition(fromCol, fromRow, 0, 0, toCol, toRow, 0, 0);
    }

    /**
     * Creates a position from a cell range.
     *
     * @param range the cell range
     * @return a new chart position
     */
    public static ChartPosition of(CellRange range) {
        return of(range.startCol(), range.startRow(), range.endCol(), range.endRow());
    }

    /**
     * Creates a position from a range string like "E1:L15".
     *
     * @param rangeRef the range reference string
     * @return a new chart position
     */
    public static ChartPosition of(String rangeRef) {
        return of(CellRange.parse(rangeRef));
    }
}
