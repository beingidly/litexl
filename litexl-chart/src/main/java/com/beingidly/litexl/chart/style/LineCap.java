package com.beingidly.litexl.chart.style;

/**
 * Line cap styles.
 */
public enum LineCap {
    /** Flat line cap. */
    FLAT("flat"),
    /** Round line cap. */
    ROUND("rnd"),
    /** Square line cap. */
    SQUARE("sq");

    private final String xmlValue;
    LineCap(String xmlValue) { this.xmlValue = xmlValue; }
    String xmlValue() { return xmlValue; }
}
