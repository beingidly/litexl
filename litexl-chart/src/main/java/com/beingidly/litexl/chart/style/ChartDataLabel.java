package com.beingidly.litexl.chart.style;

import org.jspecify.annotations.Nullable;

/**
 * Data label configuration for chart series.
 */
public record ChartDataLabel(
    boolean showValue,
    boolean showCategory,
    boolean showSeriesName,
    boolean showPercent,
    boolean showLeaderLines,
    @Nullable String separator
) {
    public static ChartDataLabel showValues() {
        return new ChartDataLabel(true, false, false, false, false, null);
    }

    public static ChartDataLabel withPercent() {
        return new ChartDataLabel(false, false, false, true, false, null);
    }

    public static Builder builder() { return new Builder(); }

    public static final class Builder {
        private boolean showValue;
        private boolean showCategory;
        private boolean showSeriesName;
        private boolean showPercent;
        private boolean showLeaderLines;
        private @Nullable String separator;

        private Builder() {}

        public Builder showValue(boolean v) { this.showValue = v; return this; }
        public Builder showCategory(boolean v) { this.showCategory = v; return this; }
        public Builder showSeriesName(boolean v) { this.showSeriesName = v; return this; }
        public Builder showPercent(boolean v) { this.showPercent = v; return this; }
        public Builder showLeaderLines(boolean v) { this.showLeaderLines = v; return this; }
        public Builder separator(String sep) { this.separator = sep; return this; }

        public ChartDataLabel build() {
            return new ChartDataLabel(showValue, showCategory, showSeriesName, showPercent, showLeaderLines, separator);
        }
    }
}
