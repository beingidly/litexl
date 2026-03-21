package com.beingidly.litexl.style;

/**
 * Represents alignment properties for a cell style.
 *
 * @param horizontal the horizontal alignment
 * @param vertical the vertical alignment
 */
public record Alignment(HAlign horizontal, VAlign vertical) {

    /**
     * Default alignment (general, bottom).
     */
    public static final Alignment DEFAULT = new Alignment(HAlign.GENERAL, VAlign.BOTTOM);

    /**
     * Creates an alignment with the given horizontal and vertical values.
     *
     * @param h the horizontal alignment
     * @param v the vertical alignment
     * @return a new alignment instance
     */
    public static Alignment of(HAlign h, VAlign v) {
        return new Alignment(h, v);
    }

    /**
     * Creates a center-center alignment.
     *
     * @return a center-center alignment instance
     */
    public static Alignment center() {
        return new Alignment(HAlign.CENTER, VAlign.MIDDLE);
    }

    /**
     * Creates a left-top alignment.
     *
     * @return a left-top alignment instance
     */
    public static Alignment leftTop() {
        return new Alignment(HAlign.LEFT, VAlign.TOP);
    }
}
