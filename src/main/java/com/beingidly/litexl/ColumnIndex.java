package com.beingidly.litexl;

/**
 * Represents a column index that may or may not exist.
 *
 * <p>Use this instead of magic values like -1 to indicate "no column".</p>
 *
 * <pre>{@code
 * ColumnIndex lastCol = row.lastColumn();
 * if (lastCol.exists()) {
 *     int index = lastCol.getValue();
 * }
 * }</pre>
 */
public final class ColumnIndex {

    private static final ColumnIndex NONE = new ColumnIndex(-1);
    private final int value;

    private ColumnIndex(int value) {
        this.value = value;
    }

    /**
     * Returns a column index representing "no column".
     *
     * @return the singleton "none" column index
     */
    public static ColumnIndex none() {
        return NONE;
    }

    /**
     * Returns a column index with the specified value.
     *
     * @param index the 0-based column index (must be non-negative)
     * @return a new column index
     * @throws IllegalArgumentException if index is negative
     */
    public static ColumnIndex of(int index) {
        if (index < 0) {
            throw new IllegalArgumentException("Index must be non-negative: " + index);
        }
        return new ColumnIndex(index);
    }

    /**
     * Returns true if this represents an existing column.
     *
     * @return true if this column index exists
     */
    public boolean exists() {
        return this != NONE;
    }

    /**
     * Returns the column index value.
     *
     * @return the column index
     * @throws IllegalStateException if no column exists
     */
    public int getValue() {
        if (!exists()) {
            throw new IllegalStateException("No column exists");
        }
        return value;
    }

    @Override
    public String toString() {
        return exists() ? "ColumnIndex[" + value + "]" : "ColumnIndex[none]";
    }
}
