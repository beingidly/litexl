package com.beingidly.litexl.chart;

/**
 * Bridge for exposing package-private XML serialization values to the internal writer package.
 * This class is intentionally public but should only be used by {@code com.beingidly.litexl.chart.internal}.
 */
public final class ChartInternalAccess {

    private ChartInternalAccess() {}

    /**
     * Returns the XML tag for the given chart type.
     *
     * @param type the chart type
     * @return the XML tag string
     */
    public static String xmlTag(ChartType type) { return type.xmlTag(); }

    /**
     * Returns the XML value for the given grouping.
     *
     * @param g the grouping
     * @return the XML string representation
     */
    public static String xmlValue(Grouping g) { return g.xmlValue(); }

    /**
     * Returns the XML value for the given bar direction.
     *
     * @param d the bar direction
     * @return the XML string representation
     */
    public static String xmlValue(BarDirection d) { return d.xmlValue(); }

    /**
     * Returns the XML value for the given scatter style.
     *
     * @param s the scatter style
     * @return the XML string representation
     */
    public static String xmlValue(ScatterStyle s) { return s.xmlValue(); }

    /**
     * Returns the XML value for the given radar style.
     *
     * @param s the radar style
     * @return the XML string representation
     */
    public static String xmlValue(RadarStyle s) { return s.xmlValue(); }

    /**
     * Returns the XML value for the given display blanks mode.
     *
     * @param d the display blanks mode
     * @return the XML string representation
     */
    public static String xmlValue(DisplayBlanks d) { return d.xmlValue(); }

    /**
     * Returns the XML value for the given legend position.
     *
     * @param p the legend position
     * @return the XML string representation
     */
    public static String xmlValue(LegendPosition p) { return p.xmlValue(); }

    /**
     * Returns the XML value for the given error bar type.
     *
     * @param t the error bar type
     * @return the XML string representation
     */
    public static String xmlValue(ChartErrorBars.Type t) { return t.xmlValue(); }

    /**
     * Returns the XML value for the given error bar direction.
     *
     * @param d the error bar direction
     * @return the XML string representation
     */
    public static String xmlValue(ChartErrorBars.Direction d) { return d.xmlValue(); }

    /**
     * Returns the XML value for the given error bar value type.
     *
     * @param vt the error bar value type
     * @return the XML string representation
     */
    public static String xmlValue(ChartErrorBars.ValueType vt) { return vt.xmlValue(); }
}
