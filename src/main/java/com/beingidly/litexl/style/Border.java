package com.beingidly.litexl.style;

/**
 * Represents border properties for a cell style.
 */
public record Border(
    BorderSide left,
    BorderSide right,
    BorderSide top,
    BorderSide bottom
) {
    /**
     * Represents a single border side.
     */
    public record BorderSide(BorderStyle style, int color) {
        public static final BorderSide NONE = new BorderSide(BorderStyle.NONE, 0);

        public static BorderSide thin(int color) {
            return new BorderSide(BorderStyle.THIN, color);
        }

        public static BorderSide medium(int color) {
            return new BorderSide(BorderStyle.MEDIUM, color);
        }

        public static BorderSide thick(int color) {
            return new BorderSide(BorderStyle.THICK, color);
        }
    }

    /**
     * No borders.
     */
    public static final Border NONE = new Border(
        BorderSide.NONE,
        BorderSide.NONE,
        BorderSide.NONE,
        BorderSide.NONE
    );

    /**
     * Creates a border with the same style and color on all sides.
     */
    public static Border all(BorderStyle style, int color) {
        BorderSide side = new BorderSide(style, color);
        return new Border(side, side, side, side);
    }

    /**
     * Creates a thin black border on all sides.
     */
    public static Border thinBlack() {
        return all(BorderStyle.THIN, 0xFF000000);
    }
}
