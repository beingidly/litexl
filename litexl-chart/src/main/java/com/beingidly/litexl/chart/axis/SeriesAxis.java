package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Series axis for 3D charts.
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

    public static SeriesAxis of(int id, int crossAxisId) {
        return new SeriesAxis(id, null, AxisPosition.BOTTOM, AxisOrientation.MIN_MAX,
            AxisTickMark.OUTSIDE, AxisTickMark.NONE, true, crossAxisId);
    }
}
