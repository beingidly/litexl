package com.beingidly.litexl.chart;

/**
 * Position of the chart legend.
 */
public enum LegendPosition {
    BOTTOM, TOP, LEFT, RIGHT,
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
