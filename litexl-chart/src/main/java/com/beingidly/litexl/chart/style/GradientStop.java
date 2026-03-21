package com.beingidly.litexl.chart.style;

/**
 * A gradient stop defining a color at a position (0.0 to 1.0).
 */
public record GradientStop(double position, ChartColor color) {
    public GradientStop {
        if (position < 0.0 || position > 1.0) {
            throw new IllegalArgumentException("Position must be between 0.0 and 1.0");
        }
    }
}
