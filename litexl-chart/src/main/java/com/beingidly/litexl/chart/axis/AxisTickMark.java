package com.beingidly.litexl.chart.axis;

/**
 * Style of tick marks on an axis.
 */
public enum AxisTickMark {
    /** No tick marks. */
    NONE,
    /** Tick marks inside the plot area. */
    INSIDE,
    /** Tick marks outside the plot area. */
    OUTSIDE,
    /** Tick marks crossing the axis line. */
    CROSS;

    String xmlValue() {
        return switch (this) {
            case NONE -> "none";
            case INSIDE -> "in";
            case OUTSIDE -> "out";
            case CROSS -> "cross";
        };
    }
}
