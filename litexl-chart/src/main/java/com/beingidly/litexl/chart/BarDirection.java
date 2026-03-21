package com.beingidly.litexl.chart;

/**
 * Direction for bar charts.
 */
public enum BarDirection {
    /** Horizontal bars. */
    BAR,
    /** Vertical columns. */
    COLUMN;

    String xmlValue() {
        return switch (this) {
            case BAR -> "bar";
            case COLUMN -> "col";
        };
    }
}
