package com.beingidly.litexl.style;

/**
 * Represents border properties for a cell style.
 *
 * @param left the left border side
 * @param right the right border side
 * @param top the top border side
 * @param bottom the bottom border side
 */
public record Border(
    BorderSide left,
    BorderSide right,
    BorderSide top,
    BorderSide bottom
) {
    /**
     * Represents a single border side.
     *
     * @param style the border style
     * @param color the border color as ARGB integer
     */
    public record BorderSide(BorderStyle style, int color) {
        public static final BorderSide NONE = new BorderSide(BorderStyle.NONE, 0);

        /**
         * Creates a thin border side with the given color.
         *
         * @param color the border color as ARGB integer
         * @return a new thin border side
         */
        public static BorderSide thin(int color) {
            return new BorderSide(BorderStyle.THIN, color);
        }

        /**
         * Creates a medium border side with the given color.
         *
         * @param color the border color as ARGB integer
         * @return a new medium border side
         */
        public static BorderSide medium(int color) {
            return new BorderSide(BorderStyle.MEDIUM, color);
        }

        /**
         * Creates a thick border side with the given color.
         *
         * @param color the border color as ARGB integer
         * @return a new thick border side
         */
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
     *
     * @param style the border style
     * @param color the border color as ARGB integer
     * @return a new border with uniform sides
     */
    public static Border all(BorderStyle style, int color) {
        BorderSide side = new BorderSide(style, color);
        return new Border(side, side, side, side);
    }

    /**
     * Creates a thin black border on all sides.
     *
     * @return a new border with thin black sides
     */
    public static Border thinBlack() {
        return all(BorderStyle.THIN, 0xFF000000);
    }
}
