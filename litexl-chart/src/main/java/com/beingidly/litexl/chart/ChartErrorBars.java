package com.beingidly.litexl.chart;

/**
 * Error bars for chart series.
 */
public record ChartErrorBars(Type type, Direction direction, ValueType valueType, double value) {

    /** Error bar type. */
    public enum Type { BOTH, PLUS, MINUS }

    /** Error bar direction. */
    public enum Direction { X, Y }

    /** How the error value is interpreted. */
    public enum ValueType { FIXED, PERCENTAGE, STANDARD_DEVIATION, STANDARD_ERROR }

    /** Creates percentage error bars. */
    public static ChartErrorBars percentage(double percent) {
        return new ChartErrorBars(Type.BOTH, Direction.Y, ValueType.PERCENTAGE, percent);
    }

    /** Creates fixed value error bars. */
    public static ChartErrorBars fixed(double value) {
        return new ChartErrorBars(Type.BOTH, Direction.Y, ValueType.FIXED, value);
    }

    /** Creates standard deviation error bars. */
    public static ChartErrorBars standardDeviation(double multiplier) {
        return new ChartErrorBars(Type.BOTH, Direction.Y, ValueType.STANDARD_DEVIATION, multiplier);
    }

    /** Creates standard error bars. */
    public static ChartErrorBars standardError() {
        return new ChartErrorBars(Type.BOTH, Direction.Y, ValueType.STANDARD_ERROR, 1.0);
    }
}
