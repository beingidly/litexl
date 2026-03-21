package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Value (numeric) axis.
 */
public record ValueAxis(
    int id,
    @Nullable String title,
    AxisPosition position,
    AxisOrientation orientation,
    @Nullable String numberFormat,
    AxisTickMark majorTickMark,
    AxisTickMark minorTickMark,
    boolean visible,
    int crossAxisId,
    AxisCrosses crosses,
    AxisCrossBetween crossBetween,
    @Nullable Double minimum,
    @Nullable Double maximum,
    @Nullable Double majorUnit,
    @Nullable Double minorUnit,
    @Nullable Double logBase
) implements ChartAxis {

    public static ValueAxis of(int id, int crossAxisId) {
        return new ValueAxis(id, null, AxisPosition.LEFT, AxisOrientation.MIN_MAX,
            null, AxisTickMark.OUTSIDE, AxisTickMark.NONE, true, crossAxisId,
            AxisCrosses.AUTO_ZERO, AxisCrossBetween.BETWEEN,
            null, null, null, null, null);
    }

    public static Builder builder(int id, int crossAxisId) {
        return new Builder(id, crossAxisId);
    }

    public static final class Builder {
        private final int id;
        private final int crossAxisId;
        private @Nullable String title;
        private AxisPosition position = AxisPosition.LEFT;
        private AxisOrientation orientation = AxisOrientation.MIN_MAX;
        private @Nullable String numberFormat;
        private AxisTickMark majorTickMark = AxisTickMark.OUTSIDE;
        private AxisTickMark minorTickMark = AxisTickMark.NONE;
        private boolean visible = true;
        private AxisCrosses crosses = AxisCrosses.AUTO_ZERO;
        private AxisCrossBetween crossBetween = AxisCrossBetween.BETWEEN;
        private @Nullable Double minimum;
        private @Nullable Double maximum;
        private @Nullable Double majorUnit;
        private @Nullable Double minorUnit;
        private @Nullable Double logBase;

        private Builder(int id, int crossAxisId) {
            this.id = id;
            this.crossAxisId = crossAxisId;
        }

        public Builder title(String title) { this.title = title; return this; }
        public Builder position(AxisPosition position) { this.position = position; return this; }
        public Builder orientation(AxisOrientation orientation) { this.orientation = orientation; return this; }
        public Builder numberFormat(String format) { this.numberFormat = format; return this; }
        public Builder majorTickMark(AxisTickMark mark) { this.majorTickMark = mark; return this; }
        public Builder minorTickMark(AxisTickMark mark) { this.minorTickMark = mark; return this; }
        public Builder visible(boolean visible) { this.visible = visible; return this; }
        public Builder crosses(AxisCrosses crosses) { this.crosses = crosses; return this; }
        public Builder crossBetween(AxisCrossBetween crossBetween) { this.crossBetween = crossBetween; return this; }
        public Builder minimum(double min) { this.minimum = min; return this; }
        public Builder maximum(double max) { this.maximum = max; return this; }
        public Builder majorUnit(double unit) { this.majorUnit = unit; return this; }
        public Builder minorUnit(double unit) { this.minorUnit = unit; return this; }
        public Builder logBase(double base) { this.logBase = base; return this; }

        public ValueAxis build() {
            return new ValueAxis(id, title, position, orientation, numberFormat,
                majorTickMark, minorTickMark, visible, crossAxisId, crosses, crossBetween,
                minimum, maximum, majorUnit, minorUnit, logBase);
        }
    }
}
