package com.beingidly.litexl.chart.axis;

import org.jspecify.annotations.Nullable;

/**
 * Category axis for bar, line, area, and radar charts.
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

    public static CategoryAxis of(int id, int crossAxisId) {
        return new CategoryAxis(id, null, AxisPosition.BOTTOM, AxisOrientation.MIN_MAX,
            AxisTickMark.CROSS, AxisTickMark.NONE, true, crossAxisId, AxisCrossBetween.MID_CAT);
    }

    public static Builder builder(int id, int crossAxisId) {
        return new Builder(id, crossAxisId);
    }

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

        public Builder title(String title) { this.title = title; return this; }
        public Builder position(AxisPosition position) { this.position = position; return this; }
        public Builder orientation(AxisOrientation orientation) { this.orientation = orientation; return this; }
        public Builder majorTickMark(AxisTickMark mark) { this.majorTickMark = mark; return this; }
        public Builder minorTickMark(AxisTickMark mark) { this.minorTickMark = mark; return this; }
        public Builder visible(boolean visible) { this.visible = visible; return this; }
        public Builder crossBetween(AxisCrossBetween crossBetween) { this.crossBetween = crossBetween; return this; }

        public CategoryAxis build() {
            return new CategoryAxis(id, title, position, orientation, majorTickMark, minorTickMark, visible, crossAxisId, crossBetween);
        }
    }
}
