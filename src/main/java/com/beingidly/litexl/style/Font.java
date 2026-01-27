package com.beingidly.litexl.style;

/**
 * Represents font properties for a cell style.
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
     */
    public static Font of(String name, double size) {
        return new Font(name, size, 0xFF000000, false, false, false, false);
    }

    /**
     * Returns a copy with bold applied.
     */
    public Font withBold(boolean bold) {
        return new Font(name, size, color, bold, italic, underline, strikethrough);
    }

    /**
     * Returns a copy with italic applied.
     */
    public Font withItalic(boolean italic) {
        return new Font(name, size, color, bold, italic, underline, strikethrough);
    }

    /**
     * Returns a copy with the given color.
     */
    public Font withColor(int color) {
        return new Font(name, size, color, bold, italic, underline, strikethrough);
    }
}
