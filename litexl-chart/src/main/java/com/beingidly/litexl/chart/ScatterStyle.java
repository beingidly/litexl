package com.beingidly.litexl.chart;

/**
 * Style for scatter (XY) charts.
 */
public enum ScatterStyle {
    MARKER, LINE, LINE_MARKER, SMOOTH, SMOOTH_MARKER;

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
