package com.beingidly.litexl;

/**
 * Represents a rectangular range of cells.
 */
public record CellRange(int startRow, int startCol, int endRow, int endCol) {

    public CellRange {
        if (startRow < 0 || startCol < 0 || endRow < 0 || endCol < 0) {
            throw new IllegalArgumentException("Cell range coordinates must be non-negative");
        }
        if (startRow > endRow || startCol > endCol) {
            throw new IllegalArgumentException("Start coordinates must not exceed end coordinates");
        }
    }

    /**
     * Creates a CellRange from coordinates.
     */
    public static CellRange of(int r1, int c1, int r2, int c2) {
        return new CellRange(r1, c1, r2, c2);
    }

    /**
     * Creates a CellRange for a single cell.
     */
    public static CellRange of(int row, int col) {
        return new CellRange(row, col, row, col);
    }

    /**
     * Creates a CellRange from a string reference like "A1:B10" or "A1".
     * This is a convenience alias for {@link #parse(String)}.
     */
    public static CellRange of(String ref) {
        return parse(ref);
    }

    /**
     * Parses a cell range reference like "A1:B10" or "A1".
     */
    public static CellRange parse(String ref) {
        if (ref.isEmpty()) {
            throw new IllegalArgumentException("Cell reference cannot be empty");
        }

        int colonIdx = ref.indexOf(':');
        if (colonIdx < 0) {
            // Single cell
            int[] coords = CellRefUtil.parseRef(ref);
            return new CellRange(coords[0], coords[1], coords[0], coords[1]);
        }

        String startRef = ref.substring(0, colonIdx);
        String endRef = ref.substring(colonIdx + 1);

        int[] start = CellRefUtil.parseRef(startRef);
        int[] end = CellRefUtil.parseRef(endRef);

        return new CellRange(start[0], start[1], end[0], end[1]);
    }

    /**
     * Returns the range as an Excel reference like "A1:B10".
     */
    public String toRef() {
        String startRef = CellRefUtil.toRef(startRow, startCol);
        if (startRow == endRow && startCol == endCol) {
            return startRef;
        }
        return startRef + ":" + CellRefUtil.toRef(endRow, endCol);
    }

    /**
     * Returns the range as an absolute Excel reference like "$A$1:$B$10".
     */
    public String toAbsoluteRef() {
        String startRef = CellRefUtil.toAbsoluteRef(startRow, startCol);
        if (startRow == endRow && startCol == endCol) {
            return startRef;
        }
        return startRef + ":" + CellRefUtil.toAbsoluteRef(endRow, endCol);
    }

    /**
     * Returns the number of rows in this range.
     */
    public int rowCount() {
        return endRow - startRow + 1;
    }

    /**
     * Returns the number of columns in this range.
     */
    public int colCount() {
        return endCol - startCol + 1;
    }

    /**
     * Checks if the given cell is within this range.
     */
    public boolean contains(int row, int col) {
        return row >= startRow && row <= endRow && col >= startCol && col <= endCol;
    }

    @Override
    public String toString() {
        return toRef();
    }
}
