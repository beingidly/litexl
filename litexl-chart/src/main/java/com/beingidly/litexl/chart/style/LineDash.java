package com.beingidly.litexl.chart.style;

/**
 * Line dash styles.
 */
public enum LineDash {
    /** Solid line. */
    SOLID("solid"),
    /** Dashed line. */
    DASH("dash"),
    /** Dotted line. */
    DOT("dot"),
    /** Dash-dot pattern. */
    DASH_DOT("dashDot"),
    /** Long dash pattern. */
    LONG_DASH("lgDash"),
    /** Long dash-dot pattern. */
    LONG_DASH_DOT("lgDashDot"),
    /** Long dash-dot-dot pattern. */
    LONG_DASH_DOT_DOT("lgDashDotDot"),
    /** System dash pattern. */
    SYSTEM_DASH("sysDash"),
    /** System dot pattern. */
    SYSTEM_DOT("sysDot"),
    /** System dash-dot pattern. */
    SYSTEM_DASH_DOT("sysDashDot");

    private final String xmlValue;
    LineDash(String xmlValue) { this.xmlValue = xmlValue; }
    String xmlValue() { return xmlValue; }
}
