package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Date axis.
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

    public static DateAxis of(int id, int crossAxisId) {
        return new DateAxis(id, null, AxisPosition.BOTTOM, AxisOrientation.MIN_MAX,
            null, AxisTickMark.OUTSIDE, AxisTickMark.NONE, true, crossAxisId);
    }
}
