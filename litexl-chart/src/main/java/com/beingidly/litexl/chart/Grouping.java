package com.beingidly.litexl.chart;

/**
 * Grouping mode for bar, column, line, and area charts.
 */
public enum Grouping {
    STANDARD, CLUSTERED, STACKED, PERCENT_STACKED;

    String xmlValue() {
        return switch (this) {
            case STANDARD -> "standard";
            case CLUSTERED -> "clustered";
            case STACKED -> "stacked";
            case PERCENT_STACKED -> "percentStacked";
        };
    }
}
