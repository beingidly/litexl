package com.beingidly.litexl.chart.style;

import org.jspecify.annotations.Nullable;

/**
 * Marker style for line and scatter chart data points.
 *
 * @param style marker shape style
 * @param size marker size (2 to 72)
 * @param fill marker fill, or {@code null} for default
 * @param line marker outline line, or {@code null} for default
 */
public record ChartMarker(MarkerStyle style, int size, @Nullable ChartFill fill, @Nullable ChartLine line) {

    /** Validates the marker size. */
    public ChartMarker {
        if (size < 2 || size > 72) {
            throw new IllegalArgumentException("Marker size must be between 2 and 72");
        }
    }

    /**
     * Creates a marker with style and size.
     *
     * @param style the marker shape style
     * @param size the marker size
     * @return a new chart marker
     */
    public static ChartMarker of(MarkerStyle style, int size) {
        return new ChartMarker(style, size, null, null);
    }

    /**
     * Creates a marker with style and default size.
     *
     * @param style the marker shape style
     * @return a new chart marker with default size
     */
    public static ChartMarker of(MarkerStyle style) {
        return new ChartMarker(style, 5, null, null);
    }
}
