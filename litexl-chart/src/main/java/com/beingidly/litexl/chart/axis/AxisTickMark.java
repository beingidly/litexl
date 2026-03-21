package com.beingidly.litexl.chart.axis;

public enum AxisTickMark {
    NONE, INSIDE, OUTSIDE, CROSS;

    String xmlValue() {
        return switch (this) {
            case NONE -> "none";
            case INSIDE -> "in";
            case OUTSIDE -> "out";
            case CROSS -> "cross";
        };
    }
}
