package com.beingidly.litexl.chart.style;

/**
 * Pattern types for pattern fills.
 */
public enum PatternType {
    PERCENT_5("pct5"), PERCENT_10("pct10"), PERCENT_20("pct20"), PERCENT_25("pct25"),
    PERCENT_30("pct30"), PERCENT_40("pct40"), PERCENT_50("pct50"), PERCENT_60("pct60"),
    PERCENT_70("pct70"), PERCENT_75("pct75"), PERCENT_80("pct80"), PERCENT_90("pct90"),
    HORIZONTAL("horz"), VERTICAL("vert"), LIGHT_HORIZONTAL("ltHorz"), LIGHT_VERTICAL("ltVert"),
    DARK_HORIZONTAL("dkHorz"), DARK_VERTICAL("dkVert"),
    NARROW_HORIZONTAL("narHorz"), NARROW_VERTICAL("narVert"),
    DASHED_HORIZONTAL("dashHorz"), DASHED_VERTICAL("dashVert"),
    CROSS("cross"), DOWNWARD_DIAGONAL("dnDiag"), UPWARD_DIAGONAL("upDiag"),
    LIGHT_DOWNWARD_DIAGONAL("ltDnDiag"), LIGHT_UPWARD_DIAGONAL("ltUpDiag"),
    DARK_DOWNWARD_DIAGONAL("dkDnDiag"), DARK_UPWARD_DIAGONAL("dkUpDiag"),
    WIDE_DOWNWARD_DIAGONAL("wdDnDiag"), WIDE_UPWARD_DIAGONAL("wdUpDiag"),
    DASHED_DOWNWARD_DIAGONAL("dashDnDiag"), DASHED_UPWARD_DIAGONAL("dashUpDiag"),
    DIAGONAL_CROSS("diagCross"), SMALL_CHECKER("smCheck"), LARGE_CHECKER("lgCheck"),
    SMALL_GRID("smGrid"), LARGE_GRID("lgGrid"),
    DOTTED_GRID("dotGrid"), SMALL_CONFETTI("smConfetti"), LARGE_CONFETTI("lgConfetti"),
    HORIZONTAL_BRICK("horzBrick"), DIAGONAL_BRICK("diagBrick"),
    SOLID_DIAMOND("solidDmnd"), OPEN_DIAMOND("openDmnd"),
    DOTTED_DIAMOND("dotDmnd"), PLAID("plaid"), SPHERE("sphere"),
    WEAVE("weave"), DIVOT("divot"), SHINGLE("shingle"),
    WAVE("wave"), TRELLIS("trellis"), ZIG_ZAG("zigZag");

    private final String xmlValue;

    PatternType(String xmlValue) { this.xmlValue = xmlValue; }
    String xmlValue() { return xmlValue; }
}
