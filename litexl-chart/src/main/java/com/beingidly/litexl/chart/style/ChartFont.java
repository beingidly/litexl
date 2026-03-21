package com.beingidly.litexl.chart.style;

import org.jspecify.annotations.Nullable;

/**
 * Font properties for chart text.
 */
public record ChartFont(
    @Nullable String name,
    double size,
    boolean bold,
    boolean italic,
    @Nullable ChartColor color
) {
    public static ChartFont of(String name, double size) {
        return new ChartFont(name, size, false, false, null);
    }

    public static Builder builder() { return new Builder(); }

    public static final class Builder {
        private @Nullable String name;
        private double size = 10.0;
        private boolean bold;
        private boolean italic;
        private @Nullable ChartColor color;

        private Builder() {}

        public Builder name(String name) { this.name = name; return this; }
        public Builder size(double size) { this.size = size; return this; }
        public Builder bold(boolean bold) { this.bold = bold; return this; }
        public Builder italic(boolean italic) { this.italic = italic; return this; }
        public Builder color(ChartColor color) { this.color = color; return this; }

        public ChartFont build() {
            return new ChartFont(name, size, bold, italic, color);
        }
    }
}
