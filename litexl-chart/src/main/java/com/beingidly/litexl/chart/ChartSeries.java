package com.beingidly.litexl.chart;

import com.beingidly.litexl.chart.style.*;
import org.jspecify.annotations.Nullable;

import java.util.Objects;

/**
 * A data series in a chart.
 *
 * @param name series name, or {@code null} for unnamed
 * @param categories category data source, or {@code null} if not applicable
 * @param values data source for the series values
 * @param fill fill style, or {@code null} for default
 * @param line line style, or {@code null} for default
 * @param marker marker style, or {@code null} for no markers
 * @param dataLabel data label configuration, or {@code null} for no labels
 * @param errorBars error bar configuration, or {@code null} for no error bars
 * @param smooth whether to smooth the line connecting data points
 * @param explosion pie slice explosion percentage (0 for no explosion)
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
     *
     * @param name the series name
     * @param categories the category range reference
     * @param values the values range reference
     * @return a new chart series
     */
    public static ChartSeries of(String name, String categories, String values) {
        return new ChartSeries(name,
            ChartDataSource.ofRange(categories),
            ChartDataSource.ofRange(values),
            null, null, null, null, null, false, 0);
    }

    /**
     * Creates a series with values only.
     *
     * @param values the values range reference
     * @return a new chart series
     */
    public static ChartSeries of(String values) {
        return new ChartSeries(null, null,
            ChartDataSource.ofRange(values),
            null, null, null, null, null, false, 0);
    }

    /**
     * Creates a new builder.
     *
     * @return a new builder
     */
    public static Builder builder() {
        return new Builder();
    }

    /** Builder for {@link ChartSeries}. */
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

        /**
         * Sets the series name.
         *
         * @param name the series name
         * @return this builder
         */
        public Builder name(String name) { this.name = name; return this; }

        /**
         * Sets the category data source from a range reference.
         *
         * @param rangeRef the category range reference
         * @return this builder
         */
        public Builder categories(String rangeRef) { this.categories = ChartDataSource.ofRange(rangeRef); return this; }

        /**
         * Sets the category data source.
         *
         * @param src the category data source
         * @return this builder
         */
        public Builder categories(ChartDataSource src) { this.categories = src; return this; }

        /**
         * Sets the values data source from a range reference.
         *
         * @param rangeRef the values range reference
         * @return this builder
         */
        public Builder values(String rangeRef) { this.values = ChartDataSource.ofRange(rangeRef); return this; }

        /**
         * Sets the values data source.
         *
         * @param src the values data source
         * @return this builder
         */
        public Builder values(ChartDataSource src) { this.values = src; return this; }

        /**
         * Sets the fill style.
         *
         * @param fill the fill style
         * @return this builder
         */
        public Builder fill(ChartFill fill) { this.fill = fill; return this; }

        /**
         * Sets the line style.
         *
         * @param line the line style
         * @return this builder
         */
        public Builder line(ChartLine line) { this.line = line; return this; }

        /**
         * Sets the marker style.
         *
         * @param marker the marker style
         * @return this builder
         */
        public Builder marker(ChartMarker marker) { this.marker = marker; return this; }

        /**
         * Sets the marker style and size.
         *
         * @param style the marker shape style
         * @param size the marker size
         * @return this builder
         */
        public Builder marker(MarkerStyle style, int size) { this.marker = ChartMarker.of(style, size); return this; }

        /**
         * Sets the data label configuration.
         *
         * @param label the data label
         * @return this builder
         */
        public Builder dataLabel(ChartDataLabel label) { this.dataLabel = label; return this; }

        /**
         * Sets the error bar configuration.
         *
         * @param bars the error bars
         * @return this builder
         */
        public Builder errorBars(ChartErrorBars bars) { this.errorBars = bars; return this; }

        /**
         * Sets whether to smooth the line.
         *
         * @param smooth true to smooth the line
         * @return this builder
         */
        public Builder smooth(boolean smooth) { this.smooth = smooth; return this; }

        /**
         * Sets the pie slice explosion percentage.
         *
         * @param percent the explosion percentage
         * @return this builder
         */
        public Builder explosion(int percent) { this.explosion = percent; return this; }

        /**
         * Builds the chart series.
         *
         * @return a new {@link ChartSeries}
         */
        public ChartSeries build() {
            Objects.requireNonNull(values, "values is required");
            return new ChartSeries(name, categories, values, fill, line, marker,
                dataLabel, errorBars, smooth, explosion);
        }
    }
}
