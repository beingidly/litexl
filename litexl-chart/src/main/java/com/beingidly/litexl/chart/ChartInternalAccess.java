package com.beingidly.litexl.chart;

/**
 * Bridge for exposing package-private XML serialization values to the internal writer package.
 * This class is intentionally public but should only be used by {@code com.beingidly.litexl.chart.internal}.
 */
public final class ChartInternalAccess {

    private ChartInternalAccess() {}

    public static String xmlTag(ChartType type) { return type.xmlTag(); }
    public static String xmlValue(Grouping g) { return g.xmlValue(); }
    public static String xmlValue(BarDirection d) { return d.xmlValue(); }
    public static String xmlValue(ScatterStyle s) { return s.xmlValue(); }
    public static String xmlValue(RadarStyle s) { return s.xmlValue(); }
    public static String xmlValue(DisplayBlanks d) { return d.xmlValue(); }
    public static String xmlValue(LegendPosition p) { return p.xmlValue(); }
    public static String xmlValue(ChartErrorBars.Type t) { return t.xmlValue(); }
    public static String xmlValue(ChartErrorBars.Direction d) { return d.xmlValue(); }
    public static String xmlValue(ChartErrorBars.ValueType vt) { return vt.xmlValue(); }
}
