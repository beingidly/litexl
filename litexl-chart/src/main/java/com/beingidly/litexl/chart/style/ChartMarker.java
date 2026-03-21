package com.beingidly.litexl.chart.style;

import org.jspecify.annotations.Nullable;

/**
 * Marker style for line and scatter chart data points.
 */
public record ChartMarker(MarkerStyle style, int size, @Nullable ChartFill fill, @Nullable ChartLine line) {

    public ChartMarker {
        if (size < 2 || size > 72) {
            throw new IllegalArgumentException("Marker size must be between 2 and 72");
        }
    }

    public static ChartMarker of(MarkerStyle style, int size) {
        return new ChartMarker(style, size, null, null);
    }

    public static ChartMarker of(MarkerStyle style) {
        return new ChartMarker(style, 5, null, null);
    }
}
