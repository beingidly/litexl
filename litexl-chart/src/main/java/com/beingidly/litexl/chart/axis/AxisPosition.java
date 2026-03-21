package com.beingidly.litexl.chart.axis;

public enum AxisPosition {
    BOTTOM, TOP, LEFT, RIGHT;

    String xmlValue() {
        return switch (this) {
            case BOTTOM -> "b";
            case TOP -> "t";
            case LEFT -> "l";
            case RIGHT -> "r";
        };
    }
}
