package com.beingidly.litexl.chart.style;

/**
 * Theme colors from the workbook's theme.
 */
public enum ThemeColor {
    /** Dark 1 (typically black). */
    DK1,
    /** Light 1 (typically white). */
    LT1,
    /** Dark 2. */
    DK2,
    /** Light 2. */
    LT2,
    /** Accent color 1. */
    ACCENT1,
    /** Accent color 2. */
    ACCENT2,
    /** Accent color 3. */
    ACCENT3,
    /** Accent color 4. */
    ACCENT4,
    /** Accent color 5. */
    ACCENT5,
    /** Accent color 6. */
    ACCENT6,
    /** Hyperlink color. */
    HLINK,
    /** Followed hyperlink color. */
    FOL_HLINK;

    String xmlValue() {
        return switch (this) {
            case DK1 -> "dk1";
            case LT1 -> "lt1";
            case DK2 -> "dk2";
            case LT2 -> "lt2";
            case ACCENT1 -> "accent1";
            case ACCENT2 -> "accent2";
            case ACCENT3 -> "accent3";
            case ACCENT4 -> "accent4";
            case ACCENT5 -> "accent5";
            case ACCENT6 -> "accent6";
            case HLINK -> "hlink";
            case FOL_HLINK -> "folHlink";
        };
    }
}
