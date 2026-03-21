package com.beingidly.litexl.chart;

import com.beingidly.litexl.chart.style.*;
import org.jspecify.annotations.Nullable;

import java.util.Objects;

/**
 * A data series in a chart.
 */
public record ChartSeries(
    @Nullable String name,
    @Nullable ChartDataSource categories,
    ChartDataSource values,
    @Nullable ChartFill fill,
    @Nullable ChartLine line,
    @Nullable ChartMarker marker,
    @Nullable ChartDataLabel dataLabel,
    @Nullable ChartErrorBars errorBars,
    boolean smooth,
    int explosion
) {

    /**
     * Creates a series with name, categories, and values.
     *
     * <p>Range references without a sheet name (no '!') will use
     * the owning sheet's name automatically.
     */
    public static ChartSeries of(String name, String categories, String values) {
        return new ChartSeries(name,
            ChartDataSource.ofRange(categories),
            ChartDataSource.ofRange(values),
            null, null, null, null, null, false, 0);
    }

    /** Creates a series with values only. */
    public static ChartSeries of(String values) {
        return new ChartSeries(null, null,
            ChartDataSource.ofRange(values),
            null, null, null, null, null, false, 0);
    }

    public static Builder builder() {
        return new Builder();
    }

    public static final class Builder {
        private @Nullable String name;
        private @Nullable ChartDataSource categories;
        private @Nullable ChartDataSource values;
        private @Nullable ChartFill fill;
        private @Nullable ChartLine line;
        private @Nullable ChartMarker marker;
        private @Nullable ChartDataLabel dataLabel;
        private @Nullable ChartErrorBars errorBars;
        private boolean smooth;
        private int explosion;

        private Builder() {}

        public Builder name(String name) { this.name = name; return this; }
        public Builder categories(String rangeRef) { this.categories = ChartDataSource.ofRange(rangeRef); return this; }
        public Builder categories(ChartDataSource src) { this.categories = src; return this; }
        public Builder values(String rangeRef) { this.values = ChartDataSource.ofRange(rangeRef); return this; }
        public Builder values(ChartDataSource src) { this.values = src; return this; }
        public Builder fill(ChartFill fill) { this.fill = fill; return this; }
        public Builder line(ChartLine line) { this.line = line; return this; }
        public Builder marker(ChartMarker marker) { this.marker = marker; return this; }
        public Builder marker(MarkerStyle style, int size) { this.marker = ChartMarker.of(style, size); return this; }
        public Builder dataLabel(ChartDataLabel label) { this.dataLabel = label; return this; }
        public Builder errorBars(ChartErrorBars bars) { this.errorBars = bars; return this; }
        public Builder smooth(boolean smooth) { this.smooth = smooth; return this; }
        public Builder explosion(int percent) { this.explosion = percent; return this; }

        public ChartSeries build() {
            Objects.requireNonNull(values, "values is required");
            return new ChartSeries(name, categories, values, fill, line, marker,
                dataLabel, errorBars, smooth, explosion);
        }
    }
}
