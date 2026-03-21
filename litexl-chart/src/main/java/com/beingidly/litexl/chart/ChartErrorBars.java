package com.beingidly.litexl.chart;

/**
 * Error bars for chart series.
 *
 * @param type error bar type (both, plus, or minus)
 * @param direction error bar direction (X or Y)
 * @param valueType how the error value is interpreted
 * @param value the error value
 */
public record ChartErrorBars(Type type, Direction direction, ValueType valueType, double value) {

    /** Error bar type. */
    public enum Type {
        /** Both positive and negative error bars. */
        BOTH,
        /** Positive error bars only. */
        PLUS,
        /** Negative error bars only. */
        MINUS;
        String xmlValue() {
            return switch (this) {
                case BOTH -> "both";
                case PLUS -> "plus";
                case MINUS -> "minus";
            };
        }
    }

    /** Error bar direction. */
    public enum Direction {
        /** Horizontal error bars. */
        X,
        /** Vertical error bars. */
        Y;
        String xmlValue() {
            return switch (this) {
                case X -> "x";
                case Y -> "y";
            };
        }
    }

    /** How the error value is interpreted. */
    public enum ValueType {
        /** Fixed absolute value. */
        FIXED,
        /** Percentage of the data value. */
        PERCENTAGE,
        /** Standard deviation multiplier. */
        STANDARD_DEVIATION,
        /** Standard error. */
        STANDARD_ERROR;
        String xmlValue() {
            return switch (this) {
                case FIXED -> "fixed";
                case PERCENTAGE -> "percentage";
                case STANDARD_DEVIATION -> "stdDev";
                case STANDARD_ERROR -> "stdErr";
            };
        }
    }

    /**
     * Creates percentage error bars.
     *
     * @param percent the percentage value
     * @return new percentage error bars
     */
    public static ChartErrorBars percentage(double percent) {
        return new ChartErrorBars(Type.BOTH, Direction.Y, ValueType.PERCENTAGE, percent);
    }

    /**
     * Creates fixed value error bars.
     *
     * @param value the fixed error value
     * @return new fixed error bars
     */
    public static ChartErrorBars fixed(double value) {
        return new ChartErrorBars(Type.BOTH, Direction.Y, ValueType.FIXED, value);
    }

    /**
     * Creates standard deviation error bars.
     *
     * @param multiplier the standard deviation multiplier
     * @return new standard deviation error bars
     */
    public static ChartErrorBars standardDeviation(double multiplier) {
        return new ChartErrorBars(Type.BOTH, Direction.Y, ValueType.STANDARD_DEVIATION, multiplier);
    }

    /**
     * Creates standard error bars.
     *
     * @return new standard error bars
     */
    public static ChartErrorBars standardError() {
        return new ChartErrorBars(Type.BOTH, Direction.Y, ValueType.STANDARD_ERROR, 1.0);
    }
}
