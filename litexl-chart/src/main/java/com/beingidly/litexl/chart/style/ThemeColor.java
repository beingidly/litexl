package com.beingidly.litexl.chart.style;

/**
 * Theme colors from the workbook's theme.
 */
public enum ThemeColor {
    DK1, LT1, DK2, LT2,
    ACCENT1, ACCENT2, ACCENT3, ACCENT4, ACCENT5, ACCENT6,
    HLINK, FOL_HLINK;

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
