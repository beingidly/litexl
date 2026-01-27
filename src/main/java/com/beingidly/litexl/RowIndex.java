package com.beingidly.litexl;

/**
 * Represents a row index that may or may not exist.
 *
 * <p>Use this instead of magic values like -1 to indicate "no row".</p>
 *
 * <pre>{@code
 * RowIndex lastRow = sheet.lastRow();
 * if (lastRow.exists()) {
 *     int index = lastRow.getValue();
 * }
 * }</pre>
 */
public final class RowIndex {

    private static final RowIndex NONE = new RowIndex(-1);
    private final int value;

    private RowIndex(int value) {
        this.value = value;
    }

    /**
     * Returns a row index representing "no row".
     */
    public static RowIndex none() {
        return NONE;
    }

    /**
     * Returns a row index with the specified value.
     *
     * @param index the 0-based row index (must be non-negative)
     * @throws IllegalArgumentException if index is negative
     */
    public static RowIndex of(int index) {
        if (index < 0) {
            throw new IllegalArgumentException("Index must be non-negative: " + index);
        }
        return new RowIndex(index);
    }

    /**
     * Returns true if this represents an existing row.
     */
    public boolean exists() {
        return this != NONE;
    }

    /**
     * Returns the row index value.
     *
     * @throws IllegalStateException if no row exists
     */
    public int getValue() {
        if (!exists()) {
            throw new IllegalStateException("No row exists");
        }
        return value;
    }

    @Override
    public String toString() {
        return exists() ? "RowIndex[" + value + "]" : "RowIndex[none]";
    }
}
