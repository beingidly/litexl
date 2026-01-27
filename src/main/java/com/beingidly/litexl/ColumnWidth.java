package com.beingidly.litexl;

/**
 * Represents a column width value.
 *
 * <p>Column width can be either an explicit value or auto-sized.</p>
 *
 * <pre>{@code
 * ColumnWidth width = sheet.format().getColumnWidth(0);
 * if (!width.isAuto()) {
 *     double value = width.getValue();
 * }
 * }</pre>
 */
public final class ColumnWidth {

    private static final ColumnWidth AUTO = new ColumnWidth(-1);
    private final double value;

    private ColumnWidth(double value) {
        this.value = value;
    }

    /**
     * Returns a column width representing auto-size.
     */
    public static ColumnWidth auto() {
        return AUTO;
    }

    /**
     * Returns a column width with the specified value.
     *
     * @param width the width in characters (must be positive)
     * @throws IllegalArgumentException if width is not positive
     */
    public static ColumnWidth of(double width) {
        if (width <= 0) {
            throw new IllegalArgumentException("Width must be positive: " + width);
        }
        return new ColumnWidth(width);
    }

    /**
     * Returns true if this represents auto-sized width.
     */
    public boolean isAuto() {
        return this == AUTO;
    }

    /**
     * Returns the width value.
     *
     * @throws IllegalStateException if this is auto-sized
     */
    public double getValue() {
        if (isAuto()) {
            throw new IllegalStateException("Auto width has no value");
        }
        return value;
    }

    @Override
    public String toString() {
        return isAuto() ? "ColumnWidth[auto]" : "ColumnWidth[" + value + "]";
    }
}
