package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Series axis for 3D charts.
 *
 * @param id axis identifier
 * @param title axis title, or {@code null} for no title
 * @param position axis position on the chart
 * @param orientation axis orientation (min-to-max or max-to-min)
 * @param majorTickMark major tick mark style
 * @param minorTickMark minor tick mark style
 * @param visible whether the axis is visible
 * @param crossAxisId identifier of the crossing axis
 */
public record SeriesAxis(
    int id,
    @Nullable String title,
    AxisPosition position,
    AxisOrientation orientation,
    AxisTickMark majorTickMark,
    AxisTickMark minorTickMark,
    boolean visible,
    int crossAxisId
) implements ChartAxis {

    /**
     * Creates a series axis with default settings.
     *
     * @param id axis identifier
     * @param crossAxisId identifier of the crossing axis
     * @return a new series axis with defaults
     */
    public static SeriesAxis of(int id, int crossAxisId) {
        return new SeriesAxis(id, null, AxisPosition.BOTTOM, AxisOrientation.MIN_MAX,
            AxisTickMark.OUTSIDE, AxisTickMark.NONE, true, crossAxisId);
    }
}
