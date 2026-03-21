package com.beingidly.litexl.chart.axis;

/**
 * Bridge for exposing package-private XML serialization values to the internal writer package.
 * This class is intentionally public but should only be used by {@code com.beingidly.litexl.chart.internal}.
 */
public final class AxisInternalAccess {

    private AxisInternalAccess() {}

    /**
     * Returns the XML value for the given axis position.
     *
     * @param p the axis position
     * @return the XML string representation
     */
    public static String xmlValue(AxisPosition p) { return p.xmlValue(); }

    /**
     * Returns the XML value for the given axis orientation.
     *
     * @param o the axis orientation
     * @return the XML string representation
     */
    public static String xmlValue(AxisOrientation o) { return o.xmlValue(); }

    /**
     * Returns the XML value for the given axis crosses setting.
     *
     * @param c the axis crosses setting
     * @return the XML string representation
     */
    public static String xmlValue(AxisCrosses c) { return c.xmlValue(); }

    /**
     * Returns the XML value for the given axis cross-between setting.
     *
     * @param c the axis cross-between setting
     * @return the XML string representation
     */
    public static String xmlValue(AxisCrossBetween c) { return c.xmlValue(); }

    /**
     * Returns the XML value for the given tick mark style.
     *
     * @param t the tick mark style
     * @return the XML string representation
     */
    public static String xmlValue(AxisTickMark t) { return t.xmlValue(); }
}
