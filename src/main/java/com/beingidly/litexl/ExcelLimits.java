package com.beingidly.litexl;

/**
 * Excel workbook limits and validation utilities.
 *
 * <p>Excel 2007+ (.xlsx) has the following limits:</p>
 * <ul>
 *   <li>Maximum rows: 1,048,576 (2^20)</li>
 *   <li>Maximum columns: 16,384 (2^14)</li>
 * </ul>
 */
public final class ExcelLimits {

    /**
     * Maximum number of rows in an Excel worksheet.
     */
    public static final int MAX_ROWS = 1_048_576;

    /**
     * Maximum number of columns in an Excel worksheet.
     */
    public static final int MAX_COLUMNS = 16_384;

    /**
     * Maximum valid row index (0-based).
     */
    public static final int MAX_ROW_INDEX = MAX_ROWS - 1;

    /**
     * Maximum valid column index (0-based).
     */
    public static final int MAX_COLUMN_INDEX = MAX_COLUMNS - 1;

    private ExcelLimits() {}

    /**
     * Validates that a row index is within Excel's limits.
     *
     * @param row the 0-based row index
     * @throws IllegalArgumentException if the row index is out of range
     */
    public static void validateRowIndex(int row) {
        if (row < 0 || row > MAX_ROW_INDEX) {
            throw new IllegalArgumentException(
                "Row index must be 0-" + MAX_ROW_INDEX + ", got: " + row);
        }
    }

    /**
     * Validates that a column index is within Excel's limits.
     *
     * @param column the 0-based column index
     * @throws IllegalArgumentException if the column index is out of range
     */
    public static void validateColumnIndex(int column) {
        if (column < 0 || column > MAX_COLUMN_INDEX) {
            throw new IllegalArgumentException(
                "Column index must be 0-" + MAX_COLUMN_INDEX + ", got: " + column);
        }
    }

    /**
     * Validates that both row and column indices are within Excel's limits.
     *
     * @param row the 0-based row index
     * @param column the 0-based column index
     * @throws IllegalArgumentException if either index is out of range
     */
    public static void validateCellIndex(int row, int column) {
        validateRowIndex(row);
        validateColumnIndex(column);
    }
}
