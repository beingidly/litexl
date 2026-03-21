package com.beingidly.litexl.chart;

/**
 * Style for scatter (XY) charts.
 */
public enum ScatterStyle {
    /** Markers only. */
    MARKER,
    /** Lines only. */
    LINE,
    /** Lines with markers. */
    LINE_MARKER,
    /** Smooth lines only. */
    SMOOTH,
    /** Smooth lines with markers. */
    SMOOTH_MARKER;

    String xmlValue() {
        return switch (this) {
            case MARKER -> "marker";
            case LINE -> "line";
            case LINE_MARKER -> "lineMarker";
            case SMOOTH -> "smooth";
            case SMOOTH_MARKER -> "smoothMarker";
        };
    }
}
