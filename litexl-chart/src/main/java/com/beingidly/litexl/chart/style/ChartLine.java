package com.beingidly.litexl.chart.style;

import org.jspecify.annotations.Nullable;

/**
 * Line properties for chart elements.
 *
 * @param color line color, or {@code null} for default
 * @param width line width in points
 * @param dash dash style
 * @param cap line cap style
 * @param join line join style
 */
public record ChartLine(
    @Nullable ChartColor color,
    double width,
    LineDash dash,
    LineCap cap,
    LineJoin join
) {
    /**
     * Creates a line with color and width.
     *
     * @param color the line color
     * @param width the line width in points
     * @return a new chart line
     */
    public static ChartLine of(ChartColor color, double width) {
        return new ChartLine(color, width, LineDash.SOLID, LineCap.FLAT, LineJoin.ROUND);
    }

    /**
     * Creates a line with an RGB hex color and width.
     *
     * @param rgbHex the hex color code
     * @param width the line width in points
     * @return a new chart line
     */
    public static ChartLine of(String rgbHex, double width) {
        return of(ChartColor.rgb(rgbHex), width);
    }

    /**
     * Creates a new builder.
     *
     * @return a new builder
     */
    public static Builder builder() { return new Builder(); }

    /** Builder for {@link ChartLine}. */
    public static final class Builder {
        private @Nullable ChartColor color;
        private double width = 1.0;
        private LineDash dash = LineDash.SOLID;
        private LineCap cap = LineCap.FLAT;
        private LineJoin join = LineJoin.ROUND;

        private Builder() {}

        /**
         * Sets the line color.
         *
         * @param color the line color
         * @return this builder
         */
        public Builder color(ChartColor color) { this.color = color; return this; }

        /**
         * Sets the line color from an RGB hex string.
         *
         * @param rgbHex the hex color code
         * @return this builder
         */
        public Builder color(String rgbHex) { this.color = ChartColor.rgb(rgbHex); return this; }

        /**
         * Sets the line width in points.
         *
         * @param width the line width
         * @return this builder
         */
        public Builder width(double width) { this.width = width; return this; }

        /**
         * Sets the dash style.
         *
         * @param dash the dash style
         * @return this builder
         */
        public Builder dash(LineDash dash) { this.dash = dash; return this; }

        /**
         * Sets the line cap style.
         *
         * @param cap the cap style
         * @return this builder
         */
        public Builder cap(LineCap cap) { this.cap = cap; return this; }

        /**
         * Sets the line join style.
         *
         * @param join the join style
         * @return this builder
         */
        public Builder join(LineJoin join) { this.join = join; return this; }

        /**
         * Builds the chart line.
         *
         * @return a new {@link ChartLine}
         */
        public ChartLine build() {
            return new ChartLine(color, width, dash, cap, join);
        }
    }
}
