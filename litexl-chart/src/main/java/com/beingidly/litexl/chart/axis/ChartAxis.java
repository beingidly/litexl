package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Base type for chart axes.
 */
public sealed interface ChartAxis permits CategoryAxis, ValueAxis, DateAxis, SeriesAxis {
    /**
     * Returns the axis identifier.
     *
     * @return axis identifier
     */
    int id();

    /**
     * Returns the axis title, or {@code null} if none.
     *
     * @return axis title, or {@code null}
     */
    @Nullable String title();

    /**
     * Returns the axis position on the chart.
     *
     * @return axis position
     */
    AxisPosition position();

    /**
     * Returns the axis orientation.
     *
     * @return axis orientation
     */
    AxisOrientation orientation();

    /**
     * Returns the major tick mark style.
     *
     * @return major tick mark style
     */
    AxisTickMark majorTickMark();

    /**
     * Returns the minor tick mark style.
     *
     * @return minor tick mark style
     */
    AxisTickMark minorTickMark();

    /**
     * Returns whether the axis is visible.
     *
     * @return true if the axis is visible
     */
    boolean visible();

    /**
     * Returns the identifier of the crossing axis.
     *
     * @return crossing axis identifier
     */
    int crossAxisId();
}
