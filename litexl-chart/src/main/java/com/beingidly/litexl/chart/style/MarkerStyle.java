package com.beingidly.litexl.chart.style;

/**
 * Marker shape styles for chart data points.
 */
public enum MarkerStyle {
    /** Circle marker. */
    CIRCLE("circle"),
    /** Diamond marker. */
    DIAMOND("diamond"),
    /** Square marker. */
    SQUARE("square"),
    /** Star marker. */
    STAR("star"),
    /** Triangle marker. */
    TRIANGLE("triangle"),
    /** X-shaped marker. */
    X("x"),
    /** Dot marker. */
    DOT("dot"),
    /** Plus marker. */
    PLUS("plus"),
    /** Dash marker. */
    DASH("dash"),
    /** Automatic marker selection. */
    AUTO("auto"),
    /** No marker. */
    NONE("none");

    private final String xmlValue;
    MarkerStyle(String xmlValue) { this.xmlValue = xmlValue; }
    String xmlValue() { return xmlValue; }
}
