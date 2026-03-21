package com.beingidly.litexl.chart.axis;

/**
 * Position of tick labels relative to the axis.
 */
public enum AxisTickLabelPosition {
    /** Labels are placed next to the axis. */
    NEXT_TO,
    /** Labels are placed at the high end of the axis. */
    HIGH,
    /** Labels are placed at the low end of the axis. */
    LOW,
    /** No tick labels are displayed. */
    NONE;

    String xmlValue() {
        return switch (this) {
            case NEXT_TO -> "nextTo";
            case HIGH -> "high";
            case LOW -> "low";
            case NONE -> "none";
        };
    }
}
