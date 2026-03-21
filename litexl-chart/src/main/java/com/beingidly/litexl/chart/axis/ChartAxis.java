package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Base type for chart axes.
 */
public sealed interface ChartAxis permits CategoryAxis, ValueAxis, DateAxis, SeriesAxis {
    int id();
    @Nullable String title();
    AxisPosition position();
    AxisOrientation orientation();
    AxisTickMark majorTickMark();
    AxisTickMark minorTickMark();
    boolean visible();
    int crossAxisId();
}
