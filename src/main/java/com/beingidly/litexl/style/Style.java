package com.beingidly.litexl.style;

import org.jspecify.annotations.Nullable;

/**
 * Represents a cell style.
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
     */
    public static Builder builder() {
        return new Builder();
    }

    public static class Builder {
        private Font font = Font.DEFAULT;
        private Border border = Border.NONE;
        private int fillColor = 0;
        private Alignment alignment = Alignment.DEFAULT;
        private String numberFormat = null;
        private boolean wrapText = false;
        private boolean locked = true;

        public Builder font(Font font) {
            this.font = font;
            return this;
        }

        public Builder font(String name, double size) {
            this.font = Font.of(name, size);
            return this;
        }

        public Builder bold(boolean bold) {
            this.font = font.withBold(bold);
            return this;
        }

        public Builder italic(boolean italic) {
            this.font = font.withItalic(italic);
            return this;
        }

        public Builder underline(boolean underline) {
            this.font = new Font(
                font.name(), font.size(), font.color(),
                font.bold(), font.italic(), underline, font.strikethrough()
            );
            return this;
        }

        public Builder color(int argb) {
            this.font = font.withColor(argb);
            return this;
        }

        public Builder border(Border border) {
            this.border = border;
            return this;
        }

        public Builder border(BorderStyle style, int color) {
            this.border = Border.all(style, color);
            return this;
        }

        public Builder borderLeft(BorderStyle style, int color) {
            this.border = new Border(
                new Border.BorderSide(style, color),
                border.right(),
                border.top(),
                border.bottom()
            );
            return this;
        }

        public Builder borderRight(BorderStyle style, int color) {
            this.border = new Border(
                border.left(),
                new Border.BorderSide(style, color),
                border.top(),
                border.bottom()
            );
            return this;
        }

        public Builder borderTop(BorderStyle style, int color) {
            this.border = new Border(
                border.left(),
                border.right(),
                new Border.BorderSide(style, color),
                border.bottom()
            );
            return this;
        }

        public Builder borderBottom(BorderStyle style, int color) {
            this.border = new Border(
                border.left(),
                border.right(),
                border.top(),
                new Border.BorderSide(style, color)
            );
            return this;
        }

        public Builder fill(int argb) {
            this.fillColor = argb;
            return this;
        }

        public Builder alignment(Alignment alignment) {
            this.alignment = alignment;
            return this;
        }

        public Builder align(HAlign h, VAlign v) {
            this.alignment = Alignment.of(h, v);
            return this;
        }

        public Builder format(@Nullable String numberFormat) {
            this.numberFormat = numberFormat;
            return this;
        }

        public Builder wrap(boolean wrap) {
            this.wrapText = wrap;
            return this;
        }

        public Builder locked(boolean locked) {
            this.locked = locked;
            return this;
        }

        public Style build() {
            return new Style(font, border, fillColor, alignment, numberFormat, wrapText, locked);
        }
    }
}
