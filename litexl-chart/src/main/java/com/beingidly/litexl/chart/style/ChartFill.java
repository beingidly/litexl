package com.beingidly.litexl.chart.style;

import java.util.List;
import java.util.Objects;

/**
 * Fill types for chart elements.
 */
public sealed interface ChartFill {

    record Solid(ChartColor color) implements ChartFill {
        public Solid { Objects.requireNonNull(color); }
    }

    record Gradient(List<GradientStop> stops, double angle) implements ChartFill {
        public Gradient {
            if (stops.size() < 2) throw new IllegalArgumentException("At least 2 gradient stops required");
            stops = List.copyOf(stops);
        }
    }

    record Pattern(PatternType type, ChartColor foreground, ChartColor background) implements ChartFill {
        public Pattern { Objects.requireNonNull(type); Objects.requireNonNull(foreground); Objects.requireNonNull(background); }
    }

    record Picture(byte[] imageData, String mimeType, PictureFillMode mode) implements ChartFill {
        public Picture {
            Objects.requireNonNull(imageData);
            Objects.requireNonNull(mimeType);
            Objects.requireNonNull(mode);
            imageData = imageData.clone();
        }

        @Override
        public byte[] imageData() { return imageData.clone(); }
    }

    record None() implements ChartFill {}

    static ChartFill solid(ChartColor color) { return new Solid(color); }
    static ChartFill solid(String rgbHex) { return new Solid(ChartColor.rgb(rgbHex)); }
    static ChartFill gradient(ChartColor from, ChartColor to) {
        return new Gradient(List.of(new GradientStop(0.0, from), new GradientStop(1.0, to)), 0.0);
    }
    static ChartFill gradient(ChartColor from, ChartColor to, double angle) {
        return new Gradient(List.of(new GradientStop(0.0, from), new GradientStop(1.0, to)), angle);
    }
    static ChartFill gradient(List<GradientStop> stops, double angle) { return new Gradient(stops, angle); }
    static ChartFill pattern(PatternType type, ChartColor fg, ChartColor bg) { return new Pattern(type, fg, bg); }
    static ChartFill picture(byte[] data, String mimeType) { return new Picture(data, mimeType, PictureFillMode.STRETCH); }
    static ChartFill picture(byte[] data, String mimeType, PictureFillMode mode) { return new Picture(data, mimeType, mode); }
    static ChartFill none() { return new None(); }
}
