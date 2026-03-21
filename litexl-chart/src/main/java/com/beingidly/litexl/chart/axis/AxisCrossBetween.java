package com.beingidly.litexl.chart.axis;

/**
 * Determines whether the value axis crosses between or on categories.
 */
public enum AxisCrossBetween {
    /** Axis crosses between categories. */
    BETWEEN,
    /** Axis crosses at the midpoint of a category. */
    MID_CAT;

    String xmlValue() {
        return switch (this) {
            case BETWEEN -> "between";
            case MID_CAT -> "midCat";
        };
    }
}
