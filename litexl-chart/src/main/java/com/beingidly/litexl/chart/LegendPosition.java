package com.beingidly.litexl.chart;

/**
 * Position of the chart legend.
 */
public enum LegendPosition {
    /** Legend at the bottom. */
    BOTTOM,
    /** Legend at the top. */
    TOP,
    /** Legend on the left. */
    LEFT,
    /** Legend on the right. */
    RIGHT,
    /** No legend displayed. */
    NONE;

    String xmlValue() {
        return switch (this) {
            case BOTTOM -> "b";
            case TOP -> "t";
            case LEFT -> "l";
            case RIGHT -> "r";
            case NONE -> "";
        };
    }
}
