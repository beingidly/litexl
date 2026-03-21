package com.beingidly.litexl.chart;

import com.beingidly.litexl.chart.axis.*;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Represents a chart to be embedded in a worksheet.
 *
 * <p>Use static factory methods for simple charts, or the builder for full control:
 * <pre>{@code
 * // Simple
 * Chart chart = Chart.of(ChartType.BAR, "Sales",
 *     ChartPosition.of("E1:L15"),
 *     List.of(ChartSeries.of("Revenue", "$A$2:$A$13", "$B$2:$B$13")));
 *
 * // Builder
 * Chart chart = Chart.bar()
 *     .title("Sales")
 *     .position(ChartPosition.of("E1:L15"))
 *     .addSeries(ChartSeries.builder()
 *         .name("Revenue")
 *         .categories("$A$2:$A$13")
 *         .values("$B$2:$B$13")
 *         .build())
 *     .build();
 * }</pre>
 */
public record Chart(
    ChartType type,
    @Nullable ChartTitle title,
    ChartPosition position,
    List<ChartSeries> series,
    @Nullable ChartLegend legend,
    ChartPlotConfig plotConfig,
    List<ChartAxis> axes
) {
    public Chart {
        Objects.requireNonNull(type);
        Objects.requireNonNull(position);
        if (series.isEmpty()) {
            throw new IllegalArgumentException("At least one series is required");
        }
        series = List.copyOf(series);
        axes = List.copyOf(axes);
    }

    // === Static Factories ===

    /** Creates a chart with title. */
    public static Chart of(ChartType type, String title,
                           ChartPosition position, List<ChartSeries> series) {
        return new Chart(type, ChartTitle.of(title), position, series,
            ChartLegend.defaults(), ChartPlotConfig.defaults(type), List.of());
    }

    /** Creates a chart without title. */
    public static Chart of(ChartType type, ChartPosition position,
                           List<ChartSeries> series) {
        return new Chart(type, null, position, series,
            null, ChartPlotConfig.defaults(type), List.of());
    }

    // === Builder ===

    public static Builder builder(ChartType type) {
        return new Builder(type);
    }

    /** Shortcut for bar chart builder. */
    public static Builder bar() { return builder(ChartType.BAR); }
    /** Shortcut for column chart builder. */
    public static Builder column() { return builder(ChartType.COLUMN); }
    /** Shortcut for line chart builder. */
    public static Builder line() { return builder(ChartType.LINE); }
    /** Shortcut for pie chart builder. */
    public static Builder pie() { return builder(ChartType.PIE); }
    /** Shortcut for scatter chart builder. */
    public static Builder scatter() { return builder(ChartType.SCATTER); }
    /** Shortcut for area chart builder. */
    public static Builder area() { return builder(ChartType.AREA); }

    public static final class Builder {
        private final ChartType type;
        private @Nullable ChartTitle title;
        private @Nullable ChartPosition position;
        private final List<ChartSeries> series = new ArrayList<>();
        private @Nullable ChartLegend legend;
        private @Nullable Grouping grouping;
        private @Nullable BarDirection barDirection;
        private @Nullable ScatterStyle scatterStyle;
        private @Nullable RadarStyle radarStyle;
        private @Nullable ChartView3D view3D;
        private DisplayBlanks displayBlanks = DisplayBlanks.GAP;
        private boolean plotVisibleOnly = true;
        private final List<ChartAxis> axes = new ArrayList<>();

        private Builder(ChartType type) {
            this.type = type;
        }

        public Builder title(String text) { this.title = ChartTitle.of(text); return this; }
        public Builder title(ChartTitle title) { this.title = title; return this; }
        public Builder position(ChartPosition pos) { this.position = pos; return this; }
        public Builder position(String rangeRef) { this.position = ChartPosition.of(rangeRef); return this; }
        public Builder addSeries(ChartSeries s) { this.series.add(s); return this; }
        public Builder series(ChartSeries... s) { this.series.addAll(List.of(s)); return this; }
        public Builder legend(ChartLegend legend) { this.legend = legend; return this; }
        public Builder legend(LegendPosition pos) { this.legend = ChartLegend.of(pos); return this; }
        public Builder grouping(Grouping g) { this.grouping = g; return this; }
        public Builder barDirection(BarDirection d) { this.barDirection = d; return this; }
        public Builder scatterStyle(ScatterStyle s) { this.scatterStyle = s; return this; }
        public Builder radarStyle(RadarStyle s) { this.radarStyle = s; return this; }
        public Builder view3D(ChartView3D v) { this.view3D = v; return this; }
        public Builder displayBlanks(DisplayBlanks d) { this.displayBlanks = d; return this; }
        public Builder plotVisibleOnly(boolean v) { this.plotVisibleOnly = v; return this; }
        public Builder categoryAxis(CategoryAxis axis) { this.axes.add(axis); return this; }
        public Builder valueAxis(ValueAxis axis) { this.axes.add(axis); return this; }
        public Builder dateAxis(DateAxis axis) { this.axes.add(axis); return this; }
        public Builder seriesAxis(SeriesAxis axis) { this.axes.add(axis); return this; }

        public Chart build() {
            Objects.requireNonNull(position, "position is required");
            if (series.isEmpty()) {
                throw new IllegalStateException("At least one series is required");
            }

            // Apply type-based defaults for unset fields
            ChartPlotConfig defaults = ChartPlotConfig.defaults(type);
            ChartPlotConfig config = new ChartPlotConfig(
                grouping != null ? grouping : defaults.grouping(),
                barDirection != null ? barDirection : defaults.barDirection(),
                scatterStyle != null ? scatterStyle : defaults.scatterStyle(),
                radarStyle != null ? radarStyle : defaults.radarStyle(),
                view3D != null ? view3D : defaults.view3D(),
                displayBlanks, plotVisibleOnly);

            return new Chart(type, title, position, series, legend, config, axes);
        }
    }
}
