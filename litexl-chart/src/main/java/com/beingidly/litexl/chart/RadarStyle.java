package com.beingidly.litexl.chart;

/**
 * Style for radar charts.
 */
public enum RadarStyle {
    /** Standard radar chart. */
    STANDARD,
    /** Radar chart with markers. */
    MARKER,
    /** Filled radar chart. */
    FILLED;

    String xmlValue() {
        return switch (this) {
            case STANDARD -> "standard";
            case MARKER -> "marker";
            case FILLED -> "filled";
        };
    }
}
