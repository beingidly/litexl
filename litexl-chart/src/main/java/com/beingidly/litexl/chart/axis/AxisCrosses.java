package com.beingidly.litexl.chart.axis;

public enum AxisCrosses {
    AUTO_ZERO, MIN, MAX;

    String xmlValue() {
        return switch (this) {
            case AUTO_ZERO -> "autoZero";
            case MIN -> "min";
            case MAX -> "max";
        };
    }
}
