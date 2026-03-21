package com.beingidly.litexl.style;

import org.jspecify.annotations.Nullable;

/**
 * Represents a cell style.
 *
 * @param font the font properties
 * @param border the border properties
 * @param fillColor the fill color as ARGB integer
 * @param alignment the cell alignment
 * @param numberFormat the number format string, or null for general
 * @param wrapText whether text wrapping is enabled
 * @param locked whether the cell is locked when the sheet is protected
 */
public record Style(
    Font font,
    Border border,
    int fillColor,
    Alignment alignment,
    @Nullable String numberFormat,
    boolean wrapText,
    boolean locked
) {
    /**
     * Default style.
     */
    public static final Style DEFAULT = new Style(
        Font.DEFAULT,
        Border.NONE,
        0,  // No fill
        Alignment.DEFAULT,
        null,
        false,
        true
    );

    /**
     * Returns a builder for creating styles.
     *
     * @return a new builder
     */
    public static Builder builder() {
        return new Builder();
    }

    /** Builder for creating Style instances. */
    public static class Builder {
        private Font font = Font.DEFAULT;
        private Border border = Border.NONE;
        private int fillColor = 0;
        private Alignment alignment = Alignment.DEFAULT;
        private String numberFormat = null;
        private boolean wrapText = false;
        private boolean locked = true;

        /** Creates a new builder with default settings. */
        public Builder() {}

        /**
         * Sets the font.
         *
         * @param font the font
         * @return this builder
         */
        public Builder font(Font font) {
            this.font = font;
            return this;
        }

        /**
         * Sets the font by name and size.
         *
         * @param name the font name
         * @param size the size
         * @return this builder
         */
        public Builder font(String name, double size) {
            this.font = Font.of(name, size);
            return this;
        }

        /**
         * Sets bold.
         *
         * @param bold true for bold
         * @return this builder
         */
        public Builder bold(boolean bold) {
            this.font = font.withBold(bold);
            return this;
        }

        /**
         * Sets italic.
         *
         * @param italic true for italic
         * @return this builder
         */
        public Builder italic(boolean italic) {
            this.font = font.withItalic(italic);
            return this;
        }

        /**
         * Sets underline.
         *
         * @param underline true for underline
         * @return this builder
         */
        public Builder underline(boolean underline) {
            this.font = new Font(
                font.name(), font.size(), font.color(),
                font.bold(), font.italic(), underline, font.strikethrough()
            );
            return this;
        }

        /**
         * Sets the font color.
         *
         * @param argb the ARGB color
         * @return this builder
         */
        public Builder color(int argb) {
            this.font = font.withColor(argb);
            return this;
        }

        /**
         * Sets the border.
         *
         * @param border the border
         * @return this builder
         */
        public Builder border(Border border) {
            this.border = border;
            return this;
        }

        /**
         * Sets uniform border.
         *
         * @param style the border style
         * @param color the ARGB color
         * @return this builder
         */
        public Builder border(BorderStyle style, int color) {
            this.border = Border.all(style, color);
            return this;
        }

        /**
         * Sets the left border.
         *
         * @param style the border style
         * @param color the ARGB color
         * @return this builder
         */
        public Builder borderLeft(BorderStyle style, int color) {
            this.border = new Border(
                new Border.BorderSide(style, color),
                border.right(),
                border.top(),
                border.bottom()
            );
            return this;
        }

        /**
         * Sets the right border.
         *
         * @param style the border style
         * @param color the ARGB color
         * @return this builder
         */
        public Builder borderRight(BorderStyle style, int color) {
            this.border = new Border(
                border.left(),
                new Border.BorderSide(style, color),
                border.top(),
                border.bottom()
            );
            return this;
        }

        /**
         * Sets the top border.
         *
         * @param style the border style
         * @param color the ARGB color
         * @return this builder
         */
        public Builder borderTop(BorderStyle style, int color) {
            this.border = new Border(
                border.left(),
                border.right(),
                new Border.BorderSide(style, color),
                border.bottom()
            );
            return this;
        }

        /**
         * Sets the bottom border.
         *
         * @param style the border style
         * @param color the ARGB color
         * @return this builder
         */
        public Builder borderBottom(BorderStyle style, int color) {
            this.border = new Border(
                border.left(),
                border.right(),
                border.top(),
                new Border.BorderSide(style, color)
            );
            return this;
        }

        /**
         * Sets the fill color.
         *
         * @param argb the ARGB fill color
         * @return this builder
         */
        public Builder fill(int argb) {
            this.fillColor = argb;
            return this;
        }

        /**
         * Sets the alignment.
         *
         * @param alignment the alignment
         * @return this builder
         */
        public Builder alignment(Alignment alignment) {
            this.alignment = alignment;
            return this;
        }

        /**
         * Sets the alignment.
         *
         * @param h the horizontal alignment
         * @param v the vertical alignment
         * @return this builder
         */
        public Builder align(HAlign h, VAlign v) {
            this.alignment = Alignment.of(h, v);
            return this;
        }

        /**
         * Sets the number format.
         *
         * @param numberFormat the format string, or null
         * @return this builder
         */
        public Builder format(@Nullable String numberFormat) {
            this.numberFormat = numberFormat;
            return this;
        }

        /**
         * Sets text wrapping.
         *
         * @param wrap true to wrap text
         * @return this builder
         */
        public Builder wrap(boolean wrap) {
            this.wrapText = wrap;
            return this;
        }

        /**
         * Sets cell locking.
         *
         * @param locked true to lock the cell
         * @return this builder
         */
        public Builder locked(boolean locked) {
            this.locked = locked;
            return this;
        }

        /**
         * Builds the style.
         *
         * @return a new Style
         */
        public Style build() {
            return new Style(font, border, fillColor, alignment, numberFormat, wrapText, locked);
        }
    }
}
