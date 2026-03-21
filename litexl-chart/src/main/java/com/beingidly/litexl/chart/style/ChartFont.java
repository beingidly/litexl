package com.beingidly.litexl.chart.style;

import org.jspecify.annotations.Nullable;

/**
 * Font properties for chart text.
 *
 * @param name font family name, or {@code null} for default
 * @param size font size in points
 * @param bold whether the font is bold
 * @param italic whether the font is italic
 * @param color font color, or {@code null} for default
 */
public record ChartFont(
    @Nullable String name,
    double size,
    boolean bold,
    boolean italic,
    @Nullable ChartColor color
) {
    /**
     * Creates a font with name and size.
     *
     * @param name the font family name
     * @param size the font size in points
     * @return a new chart font
     */
    public static ChartFont of(String name, double size) {
        return new ChartFont(name, size, false, false, null);
    }

    /**
     * Creates a new builder.
     *
     * @return a new builder
     */
    public static Builder builder() { return new Builder(); }

    /** Builder for {@link ChartFont}. */
    public static final class Builder {
        private @Nullable String name;
        private double size = 10.0;
        private boolean bold;
        private boolean italic;
        private @Nullable ChartColor color;

        private Builder() {}

        /**
         * Sets the font family name.
         *
         * @param name the font name
         * @return this builder
         */
        public Builder name(String name) { this.name = name; return this; }

        /**
         * Sets the font size in points.
         *
         * @param size the font size
         * @return this builder
         */
        public Builder size(double size) { this.size = size; return this; }

        /**
         * Sets whether the font is bold.
         *
         * @param bold true for bold
         * @return this builder
         */
        public Builder bold(boolean bold) { this.bold = bold; return this; }

        /**
         * Sets whether the font is italic.
         *
         * @param italic true for italic
         * @return this builder
         */
        public Builder italic(boolean italic) { this.italic = italic; return this; }

        /**
         * Sets the font color.
         *
         * @param color the font color
         * @return this builder
         */
        public Builder color(ChartColor color) { this.color = color; return this; }

        /**
         * Builds the chart font.
         *
         * @return a new {@link ChartFont}
         */
        public ChartFont build() {
            return new ChartFont(name, size, bold, italic, color);
        }
    }
}
