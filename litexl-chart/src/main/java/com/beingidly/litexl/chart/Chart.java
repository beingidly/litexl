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
 *
 * @param type the chart type
 * @param title chart title, or {@code null} for no title
 * @param position position of the chart on the worksheet
 * @param series data series to plot
 * @param legend legend configuration, or {@code null} for no legend
 * @param plotConfig plot area configuration
 * @param axes custom axis definitions
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
    /** Compact constructor that validates and defensively copies collections. */
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

    /**
     * Creates a chart with title.
     *
     * @param type the chart type
     * @param title the chart title text
     * @param position position of the chart on the worksheet
     * @param series data series to plot
     * @return a new chart
     */
    public static Chart of(ChartType type, String title,
                           ChartPosition position, List<ChartSeries> series) {
        return new Chart(type, ChartTitle.of(title), position, series,
            ChartLegend.defaults(), ChartPlotConfig.defaults(type), List.of());
    }

    /**
     * Creates a chart without title.
     *
     * @param type the chart type
     * @param position position of the chart on the worksheet
     * @param series data series to plot
     * @return a new chart
     */
    public static Chart of(ChartType type, ChartPosition position,
                           List<ChartSeries> series) {
        return new Chart(type, null, position, series,
            null, ChartPlotConfig.defaults(type), List.of());
    }

    // === Builder ===

    /**
     * Creates a new builder for the given chart type.
     *
     * @param type the chart type
     * @return a new builder
     */
    public static Builder builder(ChartType type) {
        return new Builder(type);
    }

    /**
     * Shortcut for bar chart builder.
     *
     * @return a new bar chart builder
     */
    public static Builder bar() { return builder(ChartType.BAR); }

    /**
     * Shortcut for column chart builder.
     *
     * @return a new column chart builder
     */
    public static Builder column() { return builder(ChartType.COLUMN); }

    /**
     * Shortcut for line chart builder.
     *
     * @return a new line chart builder
     */
    public static Builder line() { return builder(ChartType.LINE); }

    /**
     * Shortcut for pie chart builder.
     *
     * @return a new pie chart builder
     */
    public static Builder pie() { return builder(ChartType.PIE); }

    /**
     * Shortcut for scatter chart builder.
     *
     * @return a new scatter chart builder
     */
    public static Builder scatter() { return builder(ChartType.SCATTER); }

    /**
     * Shortcut for area chart builder.
     *
     * @return a new area chart builder
     */
    public static Builder area() { return builder(ChartType.AREA); }

    /** Builder for {@link Chart}. */
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

        /**
         * Sets the chart title from text.
         *
         * @param text the title text
         * @return this builder
         */
        public Builder title(String text) { this.title = ChartTitle.of(text); return this; }

        /**
         * Sets the chart title.
         *
         * @param title the chart title
         * @return this builder
         */
        public Builder title(ChartTitle title) { this.title = title; return this; }

        /**
         * Sets the chart position.
         *
         * @param pos the chart position
         * @return this builder
         */
        public Builder position(ChartPosition pos) { this.position = pos; return this; }

        /**
         * Sets the chart position from a range reference.
         *
         * @param rangeRef the cell range reference
         * @return this builder
         */
        public Builder position(String rangeRef) { this.position = ChartPosition.of(rangeRef); return this; }

        /**
         * Adds a data series.
         *
         * @param s the series to add
         * @return this builder
         */
        public Builder addSeries(ChartSeries s) { this.series.add(s); return this; }

        /**
         * Sets the data series.
         *
         * @param s the series to set
         * @return this builder
         */
        public Builder series(ChartSeries... s) { this.series.addAll(List.of(s)); return this; }

        /**
         * Sets the chart legend.
         *
         * @param legend the legend
         * @return this builder
         */
        public Builder legend(ChartLegend legend) { this.legend = legend; return this; }

        /**
         * Sets the chart legend by position.
         *
         * @param pos the legend position
         * @return this builder
         */
        public Builder legend(LegendPosition pos) { this.legend = ChartLegend.of(pos); return this; }

        /**
         * Sets the series grouping.
         *
         * @param g the grouping
         * @return this builder
         */
        public Builder grouping(Grouping g) { this.grouping = g; return this; }

        /**
         * Sets the bar direction.
         *
         * @param d the bar direction
         * @return this builder
         */
        public Builder barDirection(BarDirection d) { this.barDirection = d; return this; }

        /**
         * Sets the scatter style.
         *
         * @param s the scatter style
         * @return this builder
         */
        public Builder scatterStyle(ScatterStyle s) { this.scatterStyle = s; return this; }

        /**
         * Sets the radar style.
         *
         * @param s the radar style
         * @return this builder
         */
        public Builder radarStyle(RadarStyle s) { this.radarStyle = s; return this; }

        /**
         * Sets the 3D view configuration.
         *
         * @param v the 3D view
         * @return this builder
         */
        public Builder view3D(ChartView3D v) { this.view3D = v; return this; }

        /**
         * Sets how blank cells are displayed.
         *
         * @param d the display blanks mode
         * @return this builder
         */
        public Builder displayBlanks(DisplayBlanks d) { this.displayBlanks = d; return this; }

        /**
         * Sets whether only visible cells are plotted.
         *
         * @param v true to plot visible cells only
         * @return this builder
         */
        public Builder plotVisibleOnly(boolean v) { this.plotVisibleOnly = v; return this; }

        /**
         * Adds a category axis.
         *
         * @param axis the category axis
         * @return this builder
         */
        public Builder categoryAxis(CategoryAxis axis) { this.axes.add(axis); return this; }

        /**
         * Adds a value axis.
         *
         * @param axis the value axis
         * @return this builder
         */
        public Builder valueAxis(ValueAxis axis) { this.axes.add(axis); return this; }

        /**
         * Adds a date axis.
         *
         * @param axis the date axis
         * @return this builder
         */
        public Builder dateAxis(DateAxis axis) { this.axes.add(axis); return this; }

        /**
         * Adds a series axis.
         *
         * @param axis the series axis
         * @return this builder
         */
        public Builder seriesAxis(SeriesAxis axis) { this.axes.add(axis); return this; }

        /**
         * Builds the chart.
         *
         * @return a new {@link Chart}
         */
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
