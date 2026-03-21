package com.beingidly.litexl.chart.axis;

public enum AxisCrossBetween {
    BETWEEN, MID_CAT;

    String xmlValue() {
        return switch (this) {
            case BETWEEN -> "between";
            case MID_CAT -> "midCat";
        };
    }
}
