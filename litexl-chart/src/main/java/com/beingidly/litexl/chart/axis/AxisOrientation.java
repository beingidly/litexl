package com.beingidly.litexl.chart.axis;

public enum AxisOrientation {
    MIN_MAX, MAX_MIN;

    String xmlValue() {
        return switch (this) {
            case MIN_MAX -> "minMax";
            case MAX_MIN -> "maxMin";
        };
    }
}
