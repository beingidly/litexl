package com.beingidly.litexl.chart.style;

import org.jspecify.annotations.Nullable;

/**
 * Data label configuration for chart series.
 *
 * @param showValue whether to display the data value
 * @param showCategory whether to display the category name
 * @param showSeriesName whether to display the series name
 * @param showPercent whether to display the percentage
 * @param showLeaderLines whether to show leader lines
 * @param separator separator between label parts, or {@code null} for default
 */
public record ChartDataLabel(
    boolean showValue,
    boolean showCategory,
    boolean showSeriesName,
    boolean showPercent,
    boolean showLeaderLines,
    @Nullable String separator
) {
    /**
     * Creates a data label that shows values.
     *
     * @return a new data label configured to show values
     */
    public static ChartDataLabel showValues() {
        return new ChartDataLabel(true, false, false, false, false, null);
    }

    /**
     * Creates a data label that shows percentages.
     *
     * @return a new data label configured to show percentages
     */
    public static ChartDataLabel withPercent() {
        return new ChartDataLabel(false, false, false, true, false, null);
    }

    /**
     * Creates a new builder.
     *
     * @return a new builder
     */
    public static Builder builder() { return new Builder(); }

    /** Builder for {@link ChartDataLabel}. */
    public static final class Builder {
        private boolean showValue;
        private boolean showCategory;
        private boolean showSeriesName;
        private boolean showPercent;
        private boolean showLeaderLines;
        private @Nullable String separator;

        private Builder() {}

        /**
         * Sets whether to show the data value.
         *
         * @param v true to show values
         * @return this builder
         */
        public Builder showValue(boolean v) { this.showValue = v; return this; }

        /**
         * Sets whether to show the category name.
         *
         * @param v true to show category names
         * @return this builder
         */
        public Builder showCategory(boolean v) { this.showCategory = v; return this; }

        /**
         * Sets whether to show the series name.
         *
         * @param v true to show series names
         * @return this builder
         */
        public Builder showSeriesName(boolean v) { this.showSeriesName = v; return this; }

        /**
         * Sets whether to show the percentage.
         *
         * @param v true to show percentages
         * @return this builder
         */
        public Builder showPercent(boolean v) { this.showPercent = v; return this; }

        /**
         * Sets whether to show leader lines.
         *
         * @param v true to show leader lines
         * @return this builder
         */
        public Builder showLeaderLines(boolean v) { this.showLeaderLines = v; return this; }

        /**
         * Sets the separator between label parts.
         *
         * @param sep the separator string
         * @return this builder
         */
        public Builder separator(String sep) { this.separator = sep; return this; }

        /**
         * Builds the data label configuration.
         *
         * @return a new {@link ChartDataLabel}
         */
        public ChartDataLabel build() {
            return new ChartDataLabel(showValue, showCategory, showSeriesName, showPercent, showLeaderLines, separator);
        }
    }
}
