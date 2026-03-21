package com.beingidly.litexl.chart;

import org.jspecify.annotations.Nullable;

/**
 * Chart-type-specific plot configuration.
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

    /** Default config for the given chart type. */
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

    public static Builder builder() {
        return new Builder();
    }

    public static final class Builder {
        private @Nullable Grouping grouping;
        private @Nullable BarDirection barDirection;
        private @Nullable ScatterStyle scatterStyle;
        private @Nullable RadarStyle radarStyle;
        private @Nullable ChartView3D view3D;
        private DisplayBlanks displayBlanks = DisplayBlanks.GAP;
        private boolean plotVisibleOnly = true;

        private Builder() {}

        public Builder grouping(Grouping grouping) { this.grouping = grouping; return this; }
        public Builder barDirection(BarDirection dir) { this.barDirection = dir; return this; }
        public Builder scatterStyle(ScatterStyle style) { this.scatterStyle = style; return this; }
        public Builder radarStyle(RadarStyle style) { this.radarStyle = style; return this; }
        public Builder view3D(ChartView3D view) { this.view3D = view; return this; }
        public Builder displayBlanks(DisplayBlanks blanks) { this.displayBlanks = blanks; return this; }
        public Builder plotVisibleOnly(boolean v) { this.plotVisibleOnly = v; return this; }

        public ChartPlotConfig build() {
            return new ChartPlotConfig(grouping, barDirection, scatterStyle, radarStyle,
                view3D, displayBlanks, plotVisibleOnly);
        }
    }
}
