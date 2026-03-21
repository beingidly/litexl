package com.beingidly.litexl.chart.style;

/**
 * Represents a color in chart styling.
 */
public sealed interface ChartColor {

    /**
     * RGB color specified as a six-character hex string.
     *
     * @param hex the hex color code (e.g. "FF0000")
     */
    record Rgb(String hex) implements ChartColor {
        /** Validates the hex color code. */
        public Rgb {
            if (hex.length() != 6 || !hex.chars().allMatch(c ->
                (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'))) {
                throw new IllegalArgumentException("Invalid hex color: " + hex);
            }
        }
    }

    /**
     * Color from the preset color palette.
     *
     * @param color the preset color
     */
    record Preset(PresetColor color) implements ChartColor {}

    /**
     * Color from the theme color palette.
     *
     * @param color the theme color
     */
    record Theme(ThemeColor color) implements ChartColor {}

    /**
     * Creates an RGB color from a hex string.
     *
     * @param hex the hex color code (e.g. "FF0000")
     * @return a new RGB chart color
     */
    static ChartColor rgb(String hex) { return new Rgb(hex); }

    /**
     * Creates an RGB color from individual components.
     *
     * @param r red component (0-255)
     * @param g green component (0-255)
     * @param b blue component (0-255)
     * @return a new RGB chart color
     */
    static ChartColor rgb(int r, int g, int b) {
        return new Rgb(String.format("%02X%02X%02X", r, g, b));
    }

    /**
     * Creates a color from a preset color.
     *
     * @param color the preset color
     * @return a new preset chart color
     */
    static ChartColor preset(PresetColor color) { return new Preset(color); }

    /**
     * Creates a color from a theme color.
     *
     * @param color the theme color
     * @return a new theme chart color
     */
    static ChartColor theme(ThemeColor color) { return new Theme(color); }
}
