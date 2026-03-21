package com.beingidly.litexl.chart;

/**
 * Chart legend configuration.
 */
public record ChartLegend(LegendPosition position, boolean overlay) {

    /** Creates a legend at the given position. */
    public static ChartLegend of(LegendPosition position) {
        return new ChartLegend(position, false);
    }

    /** Creates a default legend at the bottom. */
    public static ChartLegend defaults() {
        return new ChartLegend(LegendPosition.BOTTOM, false);
    }
}
