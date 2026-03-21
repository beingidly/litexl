package com.beingidly.litexl.chart.style;

import java.util.List;
import java.util.Objects;

/**
 * Fill types for chart elements.
 */
public sealed interface ChartFill {

    /**
     * Solid color fill.
     *
     * @param color the fill color
     */
    record Solid(ChartColor color) implements ChartFill {
        /** Validates that the color is not null. */
        public Solid { Objects.requireNonNull(color); }
    }

    /**
     * Gradient fill with multiple color stops.
     *
     * @param stops the gradient color stops
     * @param angle the gradient angle in degrees
     */
    record Gradient(List<GradientStop> stops, double angle) implements ChartFill {
        /** Validates and defensively copies the stops list. */
        public Gradient {
            if (stops.size() < 2) throw new IllegalArgumentException("At least 2 gradient stops required");
            stops = List.copyOf(stops);
        }
    }

    /**
     * Pattern fill with foreground and background colors.
     *
     * @param type the pattern type
     * @param foreground the foreground color
     * @param background the background color
     */
    record Pattern(PatternType type, ChartColor foreground, ChartColor background) implements ChartFill {
        /** Validates that all parameters are not null. */
        public Pattern { Objects.requireNonNull(type); Objects.requireNonNull(foreground); Objects.requireNonNull(background); }
    }

    /**
     * Picture fill using image data.
     *
     * @param imageData the raw image bytes
     * @param mimeType the image MIME type
     * @param mode the picture fill mode
     */
    record Picture(byte[] imageData, String mimeType, PictureFillMode mode) implements ChartFill {
        /** Validates parameters and defensively copies image data. */
        public Picture {
            Objects.requireNonNull(imageData);
            Objects.requireNonNull(mimeType);
            Objects.requireNonNull(mode);
            imageData = imageData.clone();
        }

        /**
         * Returns a defensive copy of the image data.
         *
         * @return a copy of the raw image bytes
         */
        @Override
        public byte[] imageData() { return imageData.clone(); }
    }

    /** No fill (transparent). */
    record None() implements ChartFill {}

    /**
     * Creates a solid color fill.
     *
     * @param color the fill color
     * @return a new solid fill
     */
    static ChartFill solid(ChartColor color) { return new Solid(color); }

    /**
     * Creates a solid color fill from an RGB hex string.
     *
     * @param rgbHex the hex color code
     * @return a new solid fill
     */
    static ChartFill solid(String rgbHex) { return new Solid(ChartColor.rgb(rgbHex)); }

    /**
     * Creates a linear gradient fill from two colors.
     *
     * @param from the start color
     * @param to the end color
     * @return a new gradient fill
     */
    static ChartFill gradient(ChartColor from, ChartColor to) {
        return new Gradient(List.of(new GradientStop(0.0, from), new GradientStop(1.0, to)), 0.0);
    }

    /**
     * Creates a linear gradient fill from two colors with an angle.
     *
     * @param from the start color
     * @param to the end color
     * @param angle the gradient angle in degrees
     * @return a new gradient fill
     */
    static ChartFill gradient(ChartColor from, ChartColor to, double angle) {
        return new Gradient(List.of(new GradientStop(0.0, from), new GradientStop(1.0, to)), angle);
    }

    /**
     * Creates a gradient fill from explicit stops.
     *
     * @param stops the gradient color stops
     * @param angle the gradient angle in degrees
     * @return a new gradient fill
     */
    static ChartFill gradient(List<GradientStop> stops, double angle) { return new Gradient(stops, angle); }

    /**
     * Creates a pattern fill.
     *
     * @param type the pattern type
     * @param fg the foreground color
     * @param bg the background color
     * @return a new pattern fill
     */
    static ChartFill pattern(PatternType type, ChartColor fg, ChartColor bg) { return new Pattern(type, fg, bg); }

    /**
     * Creates a stretched picture fill.
     *
     * @param data the raw image bytes
     * @param mimeType the image MIME type
     * @return a new picture fill
     */
    static ChartFill picture(byte[] data, String mimeType) { return new Picture(data, mimeType, PictureFillMode.STRETCH); }

    /**
     * Creates a picture fill with specified mode.
     *
     * @param data the raw image bytes
     * @param mimeType the image MIME type
     * @param mode the picture fill mode
     * @return a new picture fill
     */
    static ChartFill picture(byte[] data, String mimeType, PictureFillMode mode) { return new Picture(data, mimeType, mode); }

    /**
     * Creates a transparent (no fill) instance.
     *
     * @return a no-fill instance
     */
    static ChartFill none() { return new None(); }
}
