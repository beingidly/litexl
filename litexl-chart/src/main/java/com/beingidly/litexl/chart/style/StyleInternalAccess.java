package com.beingidly.litexl.chart.style;

/**
 * Bridge for exposing package-private XML serialization values to the internal writer package.
 * This class is intentionally public but should only be used by {@code com.beingidly.litexl.chart.internal}.
 */
public final class StyleInternalAccess {

    private StyleInternalAccess() {}

    public static String xmlValue(PresetColor c) { return c.xmlValue(); }
    public static String xmlValue(ThemeColor c) { return c.xmlValue(); }
    public static String xmlValue(LineDash d) { return d.xmlValue(); }
    public static String xmlValue(LineCap c) { return c.xmlValue(); }
    public static String xmlValue(PatternType t) { return t.xmlValue(); }
    public static String xmlValue(MarkerStyle s) { return s.xmlValue(); }
}
