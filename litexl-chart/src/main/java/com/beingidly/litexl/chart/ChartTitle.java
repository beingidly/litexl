package com.beingidly.litexl.chart;

import com.beingidly.litexl.chart.style.ChartFont;
import org.jspecify.annotations.Nullable;

/**
 * Chart title configuration.
 *
 * @param text the title text
 * @param font font for the title, or {@code null} for default
 * @param overlay whether the title overlaps the plot area
 */
public record ChartTitle(String text, @Nullable ChartFont font, boolean overlay) {

    /**
     * Creates a simple title.
     *
     * @param text the title text
     * @return a new chart title
     */
    public static ChartTitle of(String text) {
        return new ChartTitle(text, null, false);
    }

    /**
     * Creates a title with font.
     *
     * @param text the title text
     * @param font the font for the title
     * @return a new chart title with font
     */
    public static ChartTitle of(String text, ChartFont font) {
        return new ChartTitle(text, font, false);
    }
}
