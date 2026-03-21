package com.beingidly.litexl.chart.axis;

/**
 * Orientation of an axis (ascending or descending).
 */
public enum AxisOrientation {
    /** Values increase from minimum to maximum. */
    MIN_MAX,
    /** Values decrease from maximum to minimum. */
    MAX_MIN;

    String xmlValue() {
        return switch (this) {
            case MIN_MAX -> "minMax";
            case MAX_MIN -> "maxMin";
        };
    }
}
