package com.beingidly.litexl.chart.axis;

/**
 * Position of an axis on the chart.
 */
public enum AxisPosition {
    /** Bottom of the chart. */
    BOTTOM,
    /** Top of the chart. */
    TOP,
    /** Left side of the chart. */
    LEFT,
    /** Right side of the chart. */
    RIGHT;

    String xmlValue() {
        return switch (this) {
            case BOTTOM -> "b";
            case TOP -> "t";
            case LEFT -> "l";
            case RIGHT -> "r";
        };
    }
}
