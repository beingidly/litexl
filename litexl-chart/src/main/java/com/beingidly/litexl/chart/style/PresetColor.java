package com.beingidly.litexl.chart.style;

/**
 * Preset colors matching OOXML's {@code a:prstClr} values.
 */
public enum PresetColor {
    BLACK("black"), WHITE("white"), RED("red"), GREEN("green"), BLUE("blue"),
    YELLOW("yellow"), CYAN("cyan"), MAGENTA("magenta"),
    DARK_RED("dkRed"), DARK_GREEN("dkGreen"), DARK_BLUE("dkBlue"),
    LIGHT_GRAY("ltGray"), DARK_GRAY("dkGray"), GRAY("gray"),
    ORANGE("orange"), PINK("pink"), PURPLE("purple"),
    BROWN("brown"), OLIVE("olive"), NAVY("navy"),
    TEAL("teal"), LIME_GREEN("limeGreen"), AQUA("aqua"),
    CORAL("coral"), CRIMSON("crimson"), GOLD("gold"),
    INDIGO("indigo"), IVORY("ivory"), KHAKI("khaki"),
    LAVENDER("lavender"), MAROON("maroon"), SALMON("salmon"),
    SILVER("silver"), SKY_BLUE("skyBlue"), STEEL_BLUE("steelBlue"),
    TAN("tan"), TURQUOISE("turquoise"), VIOLET("violet"),
    WHEAT("wheat"), ALICE_BLUE("aliceBlue"), ANTIQUE_WHITE("antiqueWhite"),
    BEIGE("beige"), BISQUE("bisque"), BLANCHED_ALMOND("blanchedAlmond"),
    CHOCOLATE("chocolate"), CORNFLOWER_BLUE("cornflowerBlue"),
    DARK_CYAN("dkCyan"), DARK_GOLDENROD("dkGoldenrod"),
    DARK_MAGENTA("dkMagenta"), DARK_OLIVE_GREEN("dkOliveGreen"),
    DARK_ORANGE("dkOrange"), DARK_ORCHID("dkOrchid"),
    DARK_SALMON("dkSalmon"), DARK_SEA_GREEN("dkSeaGreen"),
    DARK_SLATE_BLUE("dkSlateBlue"), DARK_SLATE_GRAY("dkSlateGray"),
    DARK_TURQUOISE("dkTurquoise"), DARK_VIOLET("dkViolet"),
    DEEP_PINK("deepPink"), DEEP_SKY_BLUE("deepSkyBlue"),
    DIM_GRAY("dimGray"), DODGER_BLUE("dodgerBlue"),
    FIREBRICK("firebrick"), FOREST_GREEN("forestGreen"),
    GAINSBORO("gainsboro"), GHOST_WHITE("ghostWhite"),
    GOLDENROD("goldenrod"), GREEN_YELLOW("greenYellow"),
    HONEYDEW("honeydew"), HOT_PINK("hotPink"),
    INDIAN_RED("indianRed"), LAWN_GREEN("lawnGreen"),
    LEMON_CHIFFON("lemonChiffon"), LIGHT_BLUE("ltBlue"),
    LIGHT_CORAL("ltCoral"), LIGHT_CYAN("ltCyan"),
    LIGHT_GOLDENROD_YELLOW("ltGoldenrodYellow"),
    LIGHT_GREEN("ltGreen"), LIGHT_PINK("ltPink"),
    LIGHT_SALMON("ltSalmon"), LIGHT_SEA_GREEN("ltSeaGreen"),
    LIGHT_SKY_BLUE("ltSkyBlue"), LIGHT_SLATE_GRAY("ltSlateGray"),
    LIGHT_STEEL_BLUE("ltSteelBlue"), LIGHT_YELLOW("ltYellow"),
    LINEN("linen"), MEDIUM_AQUAMARINE("medAquamarine"),
    MEDIUM_BLUE("medBlue"), MEDIUM_ORCHID("medOrchid"),
    MEDIUM_PURPLE("medPurple"), MEDIUM_SEA_GREEN("medSeaGreen"),
    MEDIUM_SLATE_BLUE("medSlateBlue"), MEDIUM_SPRING_GREEN("medSpringGreen"),
    MEDIUM_TURQUOISE("medTurquoise"), MEDIUM_VIOLET_RED("medVioletRed"),
    MIDNIGHT_BLUE("midnightBlue"), MINT_CREAM("mintCream"),
    MISTY_ROSE("mistyRose"), MOCCASIN("moccasin"),
    NAVAJO_WHITE("navajoWhite"), OLD_LACE("oldLace"),
    OLIVE_DRAB("oliveDrab"), ORANGE_RED("orangeRed"),
    ORCHID("orchid"), PALE_GOLDENROD("paleGoldenrod"),
    PALE_GREEN("paleGreen"), PALE_TURQUOISE("paleTurquoise"),
    PALE_VIOLET_RED("paleVioletRed"), PAPAYA_WHIP("papayaWhip"),
    PEACH_PUFF("peachPuff"), PERU("peru"), PLUM("plum"),
    POWDER_BLUE("powderBlue"), ROSY_BROWN("rosyBrown"),
    ROYAL_BLUE("royalBlue"), SADDLE_BROWN("saddleBrown"),
    SEA_GREEN("seaGreen"), SEA_SHELL("seaShell"),
    SIENNA("sienna"), SLATE_BLUE("slateBlue"),
    SLATE_GRAY("slateGray"), SNOW("snow"),
    SPRING_GREEN("springGreen"), THISTLE("thistle"),
    TOMATO("tomato"), WHITE_SMOKE("whiteSmoke"),
    YELLOW_GREEN("yellowGreen");

    private final String xmlValue;

    PresetColor(String xmlValue) {
        this.xmlValue = xmlValue;
    }

    String xmlValue() {
        return xmlValue;
    }
}
