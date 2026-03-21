package com.beingidly.litexl.chart;

/**
 * Chart legend configuration.
 *
 * @param position legend position relative to the chart
 * @param overlay whether the legend overlaps the plot area
 */
public record ChartLegend(LegendPosition position, boolean overlay) {

    /**
     * Creates a legend at the given position.
     *
     * @param position the legend position
     * @return a new chart legend
     */
    public static ChartLegend of(LegendPosition position) {
        return new ChartLegend(position, false);
    }

    /**
     * Creates a default legend at the bottom.
     *
     * @return a new chart legend at the bottom
     */
    public static ChartLegend defaults() {
        return new ChartLegend(LegendPosition.BOTTOM, false);
    }
}
