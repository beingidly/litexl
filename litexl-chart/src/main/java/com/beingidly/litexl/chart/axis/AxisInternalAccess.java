package com.beingidly.litexl.chart.axis;

/**
 * Bridge for exposing package-private XML serialization values to the internal writer package.
 * This class is intentionally public but should only be used by {@code com.beingidly.litexl.chart.internal}.
 */
public final class AxisInternalAccess {

    private AxisInternalAccess() {}

    public static String xmlValue(AxisPosition p) { return p.xmlValue(); }
    public static String xmlValue(AxisOrientation o) { return o.xmlValue(); }
    public static String xmlValue(AxisCrosses c) { return c.xmlValue(); }
    public static String xmlValue(AxisCrossBetween c) { return c.xmlValue(); }
    public static String xmlValue(AxisTickMark t) { return t.xmlValue(); }
}
