package com.beingidly.litexl.chart.axis;

public enum AxisTickLabelPosition {
    NEXT_TO, HIGH, LOW, NONE;

    String xmlValue() {
        return switch (this) {
            case NEXT_TO -> "nextTo";
            case HIGH -> "high";
            case LOW -> "low";
            case NONE -> "none";
        };
    }
}
