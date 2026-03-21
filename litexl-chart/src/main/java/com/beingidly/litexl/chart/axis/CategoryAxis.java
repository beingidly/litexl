package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Category axis for bar, line, area, and radar charts.
 *
 * @param id axis identifier
 * @param title axis title, or {@code null} for no title
 * @param position axis position on the chart
 * @param orientation axis orientation (min-to-max or max-to-min)
 * @param majorTickMark major tick mark style
 * @param minorTickMark minor tick mark style
 * @param visible whether the axis is visible
 * @param crossAxisId identifier of the crossing axis
 * @param crossBetween whether the value axis crosses between or on categories
 */
public record CategoryAxis(
    int id,
    @Nullable String title,
    AxisPosition position,
    AxisOrientation orientation,
    AxisTickMark majorTickMark,
    AxisTickMark minorTickMark,
    boolean visible,
    int crossAxisId,
    AxisCrossBetween crossBetween
) implements ChartAxis {

    /**
     * Creates a category axis with default settings.
     *
     * @param id axis identifier
     * @param crossAxisId identifier of the crossing axis
     * @return a new category axis with defaults
     */
    public static CategoryAxis of(int id, int crossAxisId) {
        return new CategoryAxis(id, null, AxisPosition.BOTTOM, AxisOrientation.MIN_MAX,
            AxisTickMark.CROSS, AxisTickMark.NONE, true, crossAxisId, AxisCrossBetween.MID_CAT);
    }

    /**
     * Creates a builder for a category axis.
     *
     * @param id axis identifier
     * @param crossAxisId identifier of the crossing axis
     * @return a new builder
     */
    public static Builder builder(int id, int crossAxisId) {
        return new Builder(id, crossAxisId);
    }

    /** Builder for {@link CategoryAxis}. */
    public static final class Builder {
        private final int id;
        private final int crossAxisId;
        private @Nullable String title;
        private AxisPosition position = AxisPosition.BOTTOM;
        private AxisOrientation orientation = AxisOrientation.MIN_MAX;
        private AxisTickMark majorTickMark = AxisTickMark.CROSS;
        private AxisTickMark minorTickMark = AxisTickMark.NONE;
        private boolean visible = true;
        private AxisCrossBetween crossBetween = AxisCrossBetween.MID_CAT;

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
         * Sets the cross-between mode.
         *
         * @param crossBetween the cross-between setting
         * @return this builder
         */
        public Builder crossBetween(AxisCrossBetween crossBetween) { this.crossBetween = crossBetween; return this; }

        /**
         * Builds the category axis.
         *
         * @return a new {@link CategoryAxis}
         */
        public CategoryAxis build() {
            return new CategoryAxis(id, title, position, orientation, majorTickMark, minorTickMark, visible, crossAxisId, crossBetween);
        }
    }
}
