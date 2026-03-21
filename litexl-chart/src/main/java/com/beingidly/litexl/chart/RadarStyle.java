package com.beingidly.litexl.chart;

/**
 * Style for radar charts.
 */
public enum RadarStyle {
    STANDARD, MARKER, FILLED;

    String xmlValue() {
        return switch (this) {
            case STANDARD -> "standard";
            case MARKER -> "marker";
            case FILLED -> "filled";
        };
    }
}
