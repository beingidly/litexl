package com.beingidly.litexl.chart.style;

/**
 * Bridge for exposing package-private XML serialization values to the internal writer package.
 * This class is intentionally public but should only be used by {@code com.beingidly.litexl.chart.internal}.
 */
public final class StyleInternalAccess {

    private StyleInternalAccess() {}

    /**
     * Returns the XML value for the given preset color.
     *
     * @param c the preset color
     * @return the XML string representation
     */
    public static String xmlValue(PresetColor c) { return c.xmlValue(); }

    /**
     * Returns the XML value for the given theme color.
     *
     * @param c the theme color
     * @return the XML string representation
     */
    public static String xmlValue(ThemeColor c) { return c.xmlValue(); }

    /**
     * Returns the XML value for the given line dash style.
     *
     * @param d the line dash style
     * @return the XML string representation
     */
    public static String xmlValue(LineDash d) { return d.xmlValue(); }

    /**
     * Returns the XML value for the given line cap style.
     *
     * @param c the line cap style
     * @return the XML string representation
     */
    public static String xmlValue(LineCap c) { return c.xmlValue(); }

    /**
     * Returns the XML value for the given pattern type.
     *
     * @param t the pattern type
     * @return the XML string representation
     */
    public static String xmlValue(PatternType t) { return t.xmlValue(); }

    /**
     * Returns the XML value for the given marker style.
     *
     * @param s the marker style
     * @return the XML string representation
     */
    public static String xmlValue(MarkerStyle s) { return s.xmlValue(); }
}
