package com.beingidly.litexl.chart;

import com.beingidly.litexl.chart.style.ChartFont;
import org.jspecify.annotations.Nullable;

/**
 * Chart title configuration.
 */
public record ChartTitle(String text, @Nullable ChartFont font, boolean overlay) {

    /** Creates a simple title. */
    public static ChartTitle of(String text) {
        return new ChartTitle(text, null, false);
    }

    /** Creates a title with font. */
    public static ChartTitle of(String text, ChartFont font) {
        return new ChartTitle(text, font, false);
    }
}
