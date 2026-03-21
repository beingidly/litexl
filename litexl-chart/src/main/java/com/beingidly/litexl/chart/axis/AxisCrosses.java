package com.beingidly.litexl.chart.axis;

/**
 * Determines where the perpendicular axis crosses this axis.
 */
public enum AxisCrosses {
    /** Axis crosses at the automatic zero point. */
    AUTO_ZERO,
    /** Axis crosses at the minimum value. */
    MIN,
    /** Axis crosses at the maximum value. */
    MAX;

    String xmlValue() {
        return switch (this) {
            case AUTO_ZERO -> "autoZero";
            case MIN -> "min";
            case MAX -> "max";
        };
    }
}
