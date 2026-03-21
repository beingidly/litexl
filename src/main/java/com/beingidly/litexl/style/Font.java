package com.beingidly.litexl.style;

/**
 * Represents font properties for a cell style.
 *
 * @param name the font family name
 * @param size the font size in points
 * @param color the font color as ARGB integer
 * @param bold whether the font is bold
 * @param italic whether the font is italic
 * @param underline whether the font is underlined
 * @param strikethrough whether the font has strikethrough
 */
public record Font(
    String name,
    double size,
    int color,
    boolean bold,
    boolean italic,
    boolean underline,
    boolean strikethrough
) {
    /**
     * Default font (Calibri 11pt).
     */
    public static final Font DEFAULT = new Font("Calibri", 11.0, 0xFF000000, false, false, false, false);

    /**
     * Creates a simple font with name and size.
     *
     * @param name the font family name
     * @param size the font size in points
     * @return a new font
     */
    public static Font of(String name, double size) {
        return new Font(name, size, 0xFF000000, false, false, false, false);
    }

    /**
     * Returns a copy with bold applied.
     *
     * @param bold whether the font is bold
     * @return a new font with the bold setting
     */
    public Font withBold(boolean bold) {
        return new Font(name, size, color, bold, italic, underline, strikethrough);
    }

    /**
     * Returns a copy with italic applied.
     *
     * @param italic whether the font is italic
     * @return a new font with the italic setting
     */
    public Font withItalic(boolean italic) {
        return new Font(name, size, color, bold, italic, underline, strikethrough);
    }

    /**
     * Returns a copy with the given color.
     *
     * @param color the font color as ARGB integer
     * @return a new font with the color applied
     */
    public Font withColor(int color) {
        return new Font(name, size, color, bold, italic, underline, strikethrough);
    }
}
