package com.beingidly.litexl.chart;

/**
 * 3D view settings for 3D charts.
 *
 * @param xRotation   X-axis rotation angle (0-360)
 * @param yRotation   Y-axis rotation angle (0-360)
 * @param perspective perspective angle (0-240)
 */
public record ChartView3D(int xRotation, int yRotation, int perspective) {

    /**
     * Creates 3D view settings.
     *
     * @param xRotation X-axis rotation angle
     * @param yRotation Y-axis rotation angle
     * @param perspective perspective angle
     * @return a new 3D view
     */
    public static ChartView3D of(int xRotation, int yRotation, int perspective) {
        return new ChartView3D(xRotation, yRotation, perspective);
    }

    /**
     * Default 3D view.
     *
     * @return a new 3D view with default angles
     */
    public static ChartView3D defaults() {
        return new ChartView3D(15, 20, 30);
    }
}
