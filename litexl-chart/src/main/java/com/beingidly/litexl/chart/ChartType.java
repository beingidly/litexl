package com.beingidly.litexl.chart;

/**
 * Types of charts supported.
 */
public enum ChartType {
    /** Vertical bar chart. */
    BAR,
    /** Horizontal bar chart. */
    COLUMN,
    /** Line chart. */
    LINE,
    /** Pie chart. */
    PIE,
    /** Scatter (XY) chart. */
    SCATTER,
    /** Area chart. */
    AREA,
    /** Radar chart. */
    RADAR,
    /** Doughnut chart. */
    DOUGHNUT,
    /** Surface chart. */
    SURFACE,
    /** 3D bar chart. */
    BAR_3D,
    /** 3D column chart. */
    COLUMN_3D,
    /** 3D line chart. */
    LINE_3D,
    /** 3D area chart. */
    AREA_3D,
    /** 3D pie chart. */
    PIE_3D;

    /** Returns true if this is a 3D chart type. */
    public boolean is3D() {
        return name().endsWith("_3D");
    }

    /** Returns true if this chart type uses axes. */
    public boolean hasAxes() {
        return this != PIE && this != PIE_3D && this != DOUGHNUT;
    }

    /** Returns the OOXML element name for this chart type. */
    String xmlTag() {
        return switch (this) {
            case BAR -> "c:barChart";
            case COLUMN -> "c:barChart";
            case LINE -> "c:lineChart";
            case PIE -> "c:pieChart";
            case SCATTER -> "c:scatterChart";
            case AREA -> "c:areaChart";
            case RADAR -> "c:radarChart";
            case DOUGHNUT -> "c:doughnutChart";
            case SURFACE -> "c:surfaceChart";
            case BAR_3D -> "c:bar3DChart";
            case COLUMN_3D -> "c:bar3DChart";
            case LINE_3D -> "c:line3DChart";
            case AREA_3D -> "c:area3DChart";
            case PIE_3D -> "c:pie3DChart";
        };
    }
}
