package com.beingidly.litexl.chart.style;

/**
 * Pattern types for pattern fills.
 */
public enum PatternType {
    /** 5% fill. */
    PERCENT_5("pct5"),
    /** 10% fill. */
    PERCENT_10("pct10"),
    /** 20% fill. */
    PERCENT_20("pct20"),
    /** 25% fill. */
    PERCENT_25("pct25"),
    /** 30% fill. */
    PERCENT_30("pct30"),
    /** 40% fill. */
    PERCENT_40("pct40"),
    /** 50% fill. */
    PERCENT_50("pct50"),
    /** 60% fill. */
    PERCENT_60("pct60"),
    /** 70% fill. */
    PERCENT_70("pct70"),
    /** 75% fill. */
    PERCENT_75("pct75"),
    /** 80% fill. */
    PERCENT_80("pct80"),
    /** 90% fill. */
    PERCENT_90("pct90"),
    /** Horizontal lines. */
    HORIZONTAL("horz"),
    /** Vertical lines. */
    VERTICAL("vert"),
    /** Light horizontal lines. */
    LIGHT_HORIZONTAL("ltHorz"),
    /** Light vertical lines. */
    LIGHT_VERTICAL("ltVert"),
    /** Dark horizontal lines. */
    DARK_HORIZONTAL("dkHorz"),
    /** Dark vertical lines. */
    DARK_VERTICAL("dkVert"),
    /** Narrow horizontal lines. */
    NARROW_HORIZONTAL("narHorz"),
    /** Narrow vertical lines. */
    NARROW_VERTICAL("narVert"),
    /** Dashed horizontal lines. */
    DASHED_HORIZONTAL("dashHorz"),
    /** Dashed vertical lines. */
    DASHED_VERTICAL("dashVert"),
    /** Cross pattern. */
    CROSS("cross"),
    /** Downward diagonal lines. */
    DOWNWARD_DIAGONAL("dnDiag"),
    /** Upward diagonal lines. */
    UPWARD_DIAGONAL("upDiag"),
    /** Light downward diagonal lines. */
    LIGHT_DOWNWARD_DIAGONAL("ltDnDiag"),
    /** Light upward diagonal lines. */
    LIGHT_UPWARD_DIAGONAL("ltUpDiag"),
    /** Dark downward diagonal lines. */
    DARK_DOWNWARD_DIAGONAL("dkDnDiag"),
    /** Dark upward diagonal lines. */
    DARK_UPWARD_DIAGONAL("dkUpDiag"),
    /** Wide downward diagonal lines. */
    WIDE_DOWNWARD_DIAGONAL("wdDnDiag"),
    /** Wide upward diagonal lines. */
    WIDE_UPWARD_DIAGONAL("wdUpDiag"),
    /** Dashed downward diagonal lines. */
    DASHED_DOWNWARD_DIAGONAL("dashDnDiag"),
    /** Dashed upward diagonal lines. */
    DASHED_UPWARD_DIAGONAL("dashUpDiag"),
    /** Diagonal cross pattern. */
    DIAGONAL_CROSS("diagCross"),
    /** Small checker pattern. */
    SMALL_CHECKER("smCheck"),
    /** Large checker pattern. */
    LARGE_CHECKER("lgCheck"),
    /** Small grid pattern. */
    SMALL_GRID("smGrid"),
    /** Large grid pattern. */
    LARGE_GRID("lgGrid"),
    /** Dotted grid pattern. */
    DOTTED_GRID("dotGrid"),
    /** Small confetti pattern. */
    SMALL_CONFETTI("smConfetti"),
    /** Large confetti pattern. */
    LARGE_CONFETTI("lgConfetti"),
    /** Horizontal brick pattern. */
    HORIZONTAL_BRICK("horzBrick"),
    /** Diagonal brick pattern. */
    DIAGONAL_BRICK("diagBrick"),
    /** Solid diamond pattern. */
    SOLID_DIAMOND("solidDmnd"),
    /** Open diamond pattern. */
    OPEN_DIAMOND("openDmnd"),
    /** Dotted diamond pattern. */
    DOTTED_DIAMOND("dotDmnd"),
    /** Plaid pattern. */
    PLAID("plaid"),
    /** Sphere pattern. */
    SPHERE("sphere"),
    /** Weave pattern. */
    WEAVE("weave"),
    /** Divot pattern. */
    DIVOT("divot"),
    /** Shingle pattern. */
    SHINGLE("shingle"),
    /** Wave pattern. */
    WAVE("wave"),
    /** Trellis pattern. */
    TRELLIS("trellis"),
    /** Zig-zag pattern. */
    ZIG_ZAG("zigZag");

    private final String xmlValue;

    PatternType(String xmlValue) { this.xmlValue = xmlValue; }
    String xmlValue() { return xmlValue; }
}
