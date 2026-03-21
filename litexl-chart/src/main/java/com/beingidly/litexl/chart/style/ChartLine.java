package com.beingidly.litexl.chart.style;

import org.jspecify.annotations.Nullable;

/**
 * Line properties for chart elements.
 */
public record ChartLine(
    @Nullable ChartColor color,
    double width,
    LineDash dash,
    LineCap cap,
    LineJoin join
) {
    public static ChartLine of(ChartColor color, double width) {
        return new ChartLine(color, width, LineDash.SOLID, LineCap.FLAT, LineJoin.ROUND);
    }

    public static ChartLine of(String rgbHex, double width) {
        return of(ChartColor.rgb(rgbHex), width);
    }

    public static Builder builder() { return new Builder(); }

    public static final class Builder {
        private @Nullable ChartColor color;
        private double width = 1.0;
        private LineDash dash = LineDash.SOLID;
        private LineCap cap = LineCap.FLAT;
        private LineJoin join = LineJoin.ROUND;

        private Builder() {}

        public Builder color(ChartColor color) { this.color = color; return this; }
        public Builder color(String rgbHex) { this.color = ChartColor.rgb(rgbHex); return this; }
        public Builder width(double width) { this.width = width; return this; }
        public Builder dash(LineDash dash) { this.dash = dash; return this; }
        public Builder cap(LineCap cap) { this.cap = cap; return this; }
        public Builder join(LineJoin join) { this.join = join; return this; }

        public ChartLine build() {
            return new ChartLine(color, width, dash, cap, join);
        }
    }
}
