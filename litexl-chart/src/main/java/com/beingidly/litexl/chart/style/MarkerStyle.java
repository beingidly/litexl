package com.beingidly.litexl.chart.style;

public enum MarkerStyle {
    CIRCLE("circle"), DIAMOND("diamond"), SQUARE("square"), STAR("star"),
    TRIANGLE("triangle"), X("x"), DOT("dot"), PLUS("plus"), DASH("dash"),
    AUTO("auto"), NONE("none");

    private final String xmlValue;
    MarkerStyle(String xmlValue) { this.xmlValue = xmlValue; }
    String xmlValue() { return xmlValue; }
}
