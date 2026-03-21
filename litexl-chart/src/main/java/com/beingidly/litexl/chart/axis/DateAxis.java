package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Date axis.
 *
 * @param id axis identifier
 * @param title axis title, or {@code null} for no title
 * @param position axis position on the chart
 * @param orientation axis orientation (min-to-max or max-to-min)
 * @param numberFormat number format string, or {@code null} for default
 * @param majorTickMark major tick mark style
 * @param minorTickMark minor tick mark style
 * @param visible whether the axis is visible
 * @param crossAxisId identifier of the crossing axis
 */
public record DateAxis(
    int id,
    @Nullable String title,
    AxisPosition position,
    AxisOrientation orientation,
    @Nullable String numberFormat,
    AxisTickMark majorTickMark,
    AxisTickMark minorTickMark,
    boolean visible,
    int crossAxisId
) implements ChartAxis {

    /**
     * Creates a date axis with default settings.
     *
     * @param id axis identifier
     * @param crossAxisId identifier of the crossing axis
     * @return a new date axis with defaults
     */
    public static DateAxis of(int id, int crossAxisId) {
        return new DateAxis(id, null, AxisPosition.BOTTOM, AxisOrientation.MIN_MAX,
            null, AxisTickMark.OUTSIDE, AxisTickMark.NONE, true, crossAxisId);
    }
}
