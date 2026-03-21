package com.beingidly.litexl.chart.style;

public enum LineDash {
    SOLID("solid"), DASH("dash"), DOT("dot"), DASH_DOT("dashDot"),
    LONG_DASH("lgDash"), LONG_DASH_DOT("lgDashDot"), LONG_DASH_DOT_DOT("lgDashDotDot"),
    SYSTEM_DASH("sysDash"), SYSTEM_DOT("sysDot"), SYSTEM_DASH_DOT("sysDashDot");

    private final String xmlValue;
    LineDash(String xmlValue) { this.xmlValue = xmlValue; }
    String xmlValue() { return xmlValue; }
}
