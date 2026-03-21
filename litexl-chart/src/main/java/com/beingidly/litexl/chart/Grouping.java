package com.beingidly.litexl.chart;

/**
 * Grouping mode for bar, column, line, and area charts.
 */
public enum Grouping {
    /** Standard grouping (no overlap). */
    STANDARD,
    /** Clustered grouping (side-by-side). */
    CLUSTERED,
    /** Stacked grouping. */
    STACKED,
    /** Percent stacked grouping. */
    PERCENT_STACKED;

    String xmlValue() {
        return switch (this) {
            case STANDARD -> "standard";
            case CLUSTERED -> "clustered";
            case STACKED -> "stacked";
            case PERCENT_STACKED -> "percentStacked";
        };
    }
}
