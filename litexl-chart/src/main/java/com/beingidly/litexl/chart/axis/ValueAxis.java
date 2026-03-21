package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Value (numeric) axis.
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
 * @param crosses where the axis crosses the perpendicular axis
 * @param crossBetween whether the axis crosses between or on categories
 * @param minimum minimum axis value, or {@code null} for auto
 * @param maximum maximum axis value, or {@code null} for auto
 * @param majorUnit major unit spacing, or {@code null} for auto
 * @param minorUnit minor unit spacing, or {@code null} for auto
 * @param logBase logarithmic base, or {@code null} for linear scale
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

    /**
     * Creates a value axis with default settings.
     *
     * @param id axis identifier
     * @param crossAxisId identifier of the crossing axis
     * @return a new value axis with defaults
     */
    public static ValueAxis of(int id, int crossAxisId) {
        return new ValueAxis(id, null, AxisPosition.LEFT, AxisOrientation.MIN_MAX,
            null, AxisTickMark.CROSS, AxisTickMark.NONE, true, crossAxisId,
            AxisCrosses.AUTO_ZERO, AxisCrossBetween.MID_CAT,
            null, null, null, null, null);
    }

    /**
     * Creates a builder for a value axis.
     *
     * @param id axis identifier
     * @param crossAxisId identifier of the crossing axis
     * @return a new builder
     */
    public static Builder builder(int id, int crossAxisId) {
        return new Builder(id, crossAxisId);
    }

    /** Builder for {@link ValueAxis}. */
    public static final class Builder {
        private final int id;
        private final int crossAxisId;
        private @Nullable String title;
        private AxisPosition position = AxisPosition.LEFT;
        private AxisOrientation orientation = AxisOrientation.MIN_MAX;
        private @Nullable String numberFormat;
        private AxisTickMark majorTickMark = AxisTickMark.CROSS;
        private AxisTickMark minorTickMark = AxisTickMark.NONE;
        private boolean visible = true;
        private AxisCrosses crosses = AxisCrosses.AUTO_ZERO;
        private AxisCrossBetween crossBetween = AxisCrossBetween.MID_CAT;
        private @Nullable Double minimum;
        private @Nullable Double maximum;
        private @Nullable Double majorUnit;
        private @Nullable Double minorUnit;
        private @Nullable Double logBase;

        private Builder(int id, int crossAxisId) {
            this.id = id;
            this.crossAxisId = crossAxisId;
        }

        /**
         * Sets the axis title.
         *
         * @param title the title
         * @return this builder
         */
        public Builder title(String title) { this.title = title; return this; }

        /**
         * Sets the axis position.
         *
         * @param position the position
         * @return this builder
         */
        public Builder position(AxisPosition position) { this.position = position; return this; }

        /**
         * Sets the axis orientation.
         *
         * @param orientation the orientation
         * @return this builder
         */
        public Builder orientation(AxisOrientation orientation) { this.orientation = orientation; return this; }

        /**
         * Sets the number format.
         *
         * @param format the number format string
         * @return this builder
         */
        public Builder numberFormat(String format) { this.numberFormat = format; return this; }

        /**
         * Sets the major tick mark style.
         *
         * @param mark the tick mark style
         * @return this builder
         */
        public Builder majorTickMark(AxisTickMark mark) { this.majorTickMark = mark; return this; }

        /**
         * Sets the minor tick mark style.
         *
         * @param mark the tick mark style
         * @return this builder
         */
        public Builder minorTickMark(AxisTickMark mark) { this.minorTickMark = mark; return this; }

        /**
         * Sets whether the axis is visible.
         *
         * @param visible true to show the axis
         * @return this builder
         */
        public Builder visible(boolean visible) { this.visible = visible; return this; }

        /**
         * Sets where the axis crosses.
         *
         * @param crosses the crosses setting
         * @return this builder
         */
        public Builder crosses(AxisCrosses crosses) { this.crosses = crosses; return this; }

        /**
         * Sets the cross-between mode.
         *
         * @param crossBetween the cross-between setting
         * @return this builder
         */
        public Builder crossBetween(AxisCrossBetween crossBetween) { this.crossBetween = crossBetween; return this; }

        /**
         * Sets the minimum axis value.
         *
         * @param min the minimum value
         * @return this builder
         */
        public Builder minimum(double min) { this.minimum = min; return this; }

        /**
         * Sets the maximum axis value.
         *
         * @param max the maximum value
         * @return this builder
         */
        public Builder maximum(double max) { this.maximum = max; return this; }

        /**
         * Sets the major unit spacing.
         *
         * @param unit the major unit
         * @return this builder
         */
        public Builder majorUnit(double unit) { this.majorUnit = unit; return this; }

        /**
         * Sets the minor unit spacing.
         *
         * @param unit the minor unit
         * @return this builder
         */
        public Builder minorUnit(double unit) { this.minorUnit = unit; return this; }

        /**
         * Sets the logarithmic base.
         *
         * @param base the logarithmic base
         * @return this builder
         */
        public Builder logBase(double base) { this.logBase = base; return this; }

        /**
         * Builds the value axis.
         *
         * @return a new {@link ValueAxis}
         */
        public ValueAxis build() {
            return new ValueAxis(id, title, position, orientation, numberFormat,
                majorTickMark, minorTickMark, visible, crossAxisId, crosses, crossBetween,
                minimum, maximum, majorUnit, minorUnit, logBase);
        }
    }
}
