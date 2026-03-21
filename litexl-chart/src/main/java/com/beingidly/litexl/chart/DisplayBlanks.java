package com.beingidly.litexl.chart;

/**
 * How blank cells are handled in charts.
 */
public enum DisplayBlanks {
    /** Leave gaps for blank cells. */
    GAP,
    /** Span across blank cells. */
    SPAN,
    /** Treat blank cells as zero. */
    ZERO;

    String xmlValue() {
        return switch (this) {
            case GAP -> "gap";
            case SPAN -> "span";
            case ZERO -> "zero";
        };
    }
}
