package com.beingidly.litexl.chart.style;

/**
 * Represents a color in chart styling.
 */
public sealed interface ChartColor {
    record Rgb(String hex) implements ChartColor {
        public Rgb {
            if (hex.length() != 6 || !hex.chars().allMatch(c ->
                (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'))) {
                throw new IllegalArgumentException("Invalid hex color: " + hex);
            }
        }
    }

    record Preset(PresetColor color) implements ChartColor {}
    record Theme(ThemeColor color) implements ChartColor {}

    static ChartColor rgb(String hex) { return new Rgb(hex); }

    static ChartColor rgb(int r, int g, int b) {
        return new Rgb(String.format("%02X%02X%02X", r, g, b));
    }

    static ChartColor preset(PresetColor color) { return new Preset(color); }
    static ChartColor theme(ThemeColor color) { return new Theme(color); }
}
