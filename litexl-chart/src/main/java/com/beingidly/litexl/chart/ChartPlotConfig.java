package com.beingidly.litexl.chart;

import org.jspecify.annotations.Nullable;

/**
 * Chart-type-specific plot configuration.
 *
 * @param grouping series grouping mode, or {@code null} if not applicable
 * @param barDirection bar direction, or {@code null} if not a bar/column chart
 * @param scatterStyle scatter style, or {@code null} if not a scatter chart
 * @param radarStyle radar style, or {@code null} if not a radar chart
 * @param view3D 3D view settings, or {@code null} if not a 3D chart
 * @param displayBlanks how blank cells are displayed
 * @param plotVisibleOnly whether only visible cells are plotted
 */
public record ChartPlotConfig(
    @Nullable Grouping grouping,
    @Nullable BarDirection barDirection,
    @Nullable ScatterStyle scatterStyle,
    @Nullable RadarStyle radarStyle,
    @Nullable ChartView3D view3D,
    DisplayBlanks displayBlanks,
    boolean plotVisibleOnly
) {

    /**
     * Default config for the given chart type.
     *
     * @param type the chart type
     * @return a default plot configuration for the type
     */
    public static ChartPlotConfig defaults(ChartType type) {
        Grouping g = type.hasAxes() && type != ChartType.SCATTER ? Grouping.CLUSTERED : null;
        BarDirection dir = (type == ChartType.BAR || type == ChartType.BAR_3D) ? BarDirection.BAR
            : (type == ChartType.COLUMN || type == ChartType.COLUMN_3D) ? BarDirection.COLUMN
            : null;
        ScatterStyle ss = type == ChartType.SCATTER ? ScatterStyle.MARKER : null;
        RadarStyle rs = type == ChartType.RADAR ? RadarStyle.STANDARD : null;
        ChartView3D v3d = type.is3D() ? ChartView3D.defaults() : null;
        return new ChartPlotConfig(g, dir, ss, rs, v3d, DisplayBlanks.GAP, true);
    }

    /**
     * Creates a new builder.
     *
     * @return a new builder
     */
    public static Builder builder() {
        return new Builder();
    }

    /** Builder for {@link ChartPlotConfig}. */
    public static final class Builder {
        private @Nullable Grouping grouping;
        private @Nullable BarDirection barDirection;
        private @Nullable ScatterStyle scatterStyle;
        private @Nullable RadarStyle radarStyle;
        private @Nullable ChartView3D view3D;
        private DisplayBlanks displayBlanks = DisplayBlanks.GAP;
        private boolean plotVisibleOnly = true;

        private Builder() {}

        /**
         * Sets the series grouping.
         *
         * @param grouping the grouping
         * @return this builder
         */
        public Builder grouping(Grouping grouping) { this.grouping = grouping; return this; }

        /**
         * Sets the bar direction.
         *
         * @param dir the bar direction
         * @return this builder
         */
        public Builder barDirection(BarDirection dir) { this.barDirection = dir; return this; }

        /**
         * Sets the scatter style.
         *
         * @param style the scatter style
         * @return this builder
         */
        public Builder scatterStyle(ScatterStyle style) { this.scatterStyle = style; return this; }

        /**
         * Sets the radar style.
         *
         * @param style the radar style
         * @return this builder
         */
        public Builder radarStyle(RadarStyle style) { this.radarStyle = style; return this; }

        /**
         * Sets the 3D view configuration.
         *
         * @param view the 3D view
         * @return this builder
         */
        public Builder view3D(ChartView3D view) { this.view3D = view; return this; }

        /**
         * Sets how blank cells are displayed.
         *
         * @param blanks the display blanks mode
         * @return this builder
         */
        public Builder displayBlanks(DisplayBlanks blanks) { this.displayBlanks = blanks; return this; }

        /**
         * Sets whether only visible cells are plotted.
         *
         * @param v true to plot visible cells only
         * @return this builder
         */
        public Builder plotVisibleOnly(boolean v) { this.plotVisibleOnly = v; return this; }

        /**
         * Builds the plot configuration.
         *
         * @return a new {@link ChartPlotConfig}
         */
        public ChartPlotConfig build() {
            return new ChartPlotConfig(grouping, barDirection, scatterStyle, radarStyle,
                view3D, displayBlanks, plotVisibleOnly);
        }
    }
}
