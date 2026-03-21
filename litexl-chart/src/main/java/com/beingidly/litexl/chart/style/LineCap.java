package com.beingidly.litexl.chart.style;

public enum LineCap {
    FLAT("flat"), ROUND("rnd"), SQUARE("sq");

    private final String xmlValue;
    LineCap(String xmlValue) { this.xmlValue = xmlValue; }
    String xmlValue() { return xmlValue; }
}
