package com.beingidly.litexl;

import com.beingidly.litexl.style.*;

import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

/**
 * Handles reading and writing of styles.xml
 */
final class StylesXml {

    private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // Built-in number formats
    private static final Map<String, Integer> BUILTIN_FORMATS = Map.ofEntries(
        Map.entry("General", 0),
        Map.entry("0", 1),
        Map.entry("0.00", 2),
        Map.entry("#,##0", 3),
        Map.entry("#,##0.00", 4),
        Map.entry("0%", 9),
        Map.entry("0.00%", 10),
        Map.entry("0.00E+00", 11),
        Map.entry("# ?/?", 12),
        Map.entry("# ??/??", 13),
        Map.entry("mm-dd-yy", 14),
        Map.entry("d-mmm-yy", 15),
        Map.entry("d-mmm", 16),
        Map.entry("mmm-yy", 17),
        Map.entry("h:mm AM/PM", 18),
        Map.entry("h:mm:ss AM/PM", 19),
        Map.entry("h:mm", 20),
        Map.entry("h:mm:ss", 21),
        Map.entry("m/d/yy h:mm", 22),
        Map.entry("@", 49)
    );

    private StylesXml() {}

    public static void write(OutputStream os, List<Style> styles) throws IOException {
        // Collect unique components
        List<Font> fonts = new ArrayList<>();
        List<Integer> fills = new ArrayList<>();
        List<Border> borders = new ArrayList<>();
        Map<String, Integer> numFmtMap = new LinkedHashMap<>();
        int nextNumFmtId = 164; // Custom formats start at 164

        // Always add default components
        fonts.add(Font.DEFAULT);
        fills.add(0); // No fill
        fills.add(0x00808080); // Gray 125 pattern (required by Excel)
        borders.add(Border.NONE);

        // Build indices for each style
        List<int[]> xfIndices = new ArrayList<>(); // [fontId, fillId, borderId, numFmtId]

        for (Style style : styles) {
            int fontId = indexOfOrAdd(fonts, style.font());
            int fillId = indexOfFillOrAdd(fills, style.fillColor());
            int borderId = indexOfOrAdd(borders, style.border());

            int numFmtId = 0;
            if (style.numberFormat() != null && !style.numberFormat().isEmpty()) {
                Integer builtIn = BUILTIN_FORMATS.get(style.numberFormat());
                if (builtIn != null) {
                    numFmtId = builtIn;
                } else {
                    Integer existing = numFmtMap.get(style.numberFormat());
                    if (existing != null) {
                        numFmtId = existing;
                    } else {
                        numFmtId = nextNumFmtId++;
                        numFmtMap.put(style.numberFormat(), numFmtId);
                    }
                }
            }

            xfIndices.add(new int[]{fontId, fillId, borderId, numFmtId});
        }

        try (XmlWriter xml = new XmlWriter(os)) {
            xml.startDocument();
            xml.startElement("styleSheet");
            xml.attribute("xmlns", NS);

            // Number formats
            if (!numFmtMap.isEmpty()) {
                xml.startElement("numFmts");
                xml.attribute("count", String.valueOf(numFmtMap.size()));
                for (Map.Entry<String, Integer> entry : numFmtMap.entrySet()) {
                    xml.emptyElement("numFmt");
                    xml.attribute("numFmtId", String.valueOf(entry.getValue()));
                    xml.attribute("formatCode", entry.getKey());
                }
                xml.endElement();
            }

            // Fonts
            xml.startElement("fonts");
            xml.attribute("count", String.valueOf(fonts.size()));
            for (Font font : fonts) {
                writeFont(xml, font);
            }
            xml.endElement();

            // Fills
            xml.startElement("fills");
            xml.attribute("count", String.valueOf(fills.size()));
            for (int i = 0; i < fills.size(); i++) {
                xml.startElement("fill");
                if (i == 0) {
                    xml.emptyElement("patternFill");
                    xml.attribute("patternType", "none");
                } else if (i == 1) {
                    xml.emptyElement("patternFill");
                    xml.attribute("patternType", "gray125");
                } else {
                    xml.startElement("patternFill");
                    xml.attribute("patternType", "solid");
                    xml.emptyElement("fgColor");
                    xml.attribute("rgb", toArgbString(fills.get(i)));
                    xml.emptyElement("bgColor");
                    xml.attribute("indexed", "64");
                    xml.endElement();
                }
                xml.endElement();
            }
            xml.endElement();

            // Borders
            xml.startElement("borders");
            xml.attribute("count", String.valueOf(borders.size()));
            for (Border border : borders) {
                writeBorder(xml, border);
            }
            xml.endElement();

            // Cell style xfs (base styles)
            xml.startElement("cellStyleXfs");
            xml.attribute("count", "1");
            xml.emptyElement("xf");
            xml.attribute("numFmtId", "0");
            xml.attribute("fontId", "0");
            xml.attribute("fillId", "0");
            xml.attribute("borderId", "0");
            xml.endElement();

            // Cell xfs (actual cell styles)
            xml.startElement("cellXfs");
            xml.attribute("count", String.valueOf(xfIndices.size()));
            for (int i = 0; i < xfIndices.size(); i++) {
                int[] idx = xfIndices.get(i);
                Style style = styles.get(i);

                xml.startElement("xf");
                xml.attribute("numFmtId", String.valueOf(idx[3]));
                xml.attribute("fontId", String.valueOf(idx[0]));
                xml.attribute("fillId", String.valueOf(idx[1]));
                xml.attribute("borderId", String.valueOf(idx[2]));
                xml.attribute("xfId", "0");

                if (idx[0] > 0) {
                    xml.attribute("applyFont", "1");
                }
                if (idx[1] > 1) {
                    xml.attribute("applyFill", "1");
                }
                if (idx[2] > 0) {
                    xml.attribute("applyBorder", "1");
                }
                if (idx[3] > 0) {
                    xml.attribute("applyNumberFormat", "1");
                }

                // Alignment
                if (style.alignment() != null &&
                    (style.alignment().horizontal() != HAlign.GENERAL ||
                     style.alignment().vertical() != VAlign.BOTTOM ||
                     style.wrapText())) {
                    xml.attribute("applyAlignment", "1");
                    xml.emptyElement("alignment");
                    if (style.alignment().horizontal() != HAlign.GENERAL) {
                        xml.attribute("horizontal", style.alignment().horizontal().name().toLowerCase());
                    }
                    if (style.alignment().vertical() != VAlign.BOTTOM) {
                        xml.attribute("vertical", style.alignment().vertical().name().toLowerCase());
                    }
                    if (style.wrapText()) {
                        xml.attribute("wrapText", "1");
                    }
                }

                // Protection
                if (!style.locked()) {
                    xml.attribute("applyProtection", "1");
                    xml.emptyElement("protection");
                    xml.attribute("locked", "0");
                }

                xml.endElement();
            }
            xml.endElement();

            // Cell styles
            xml.startElement("cellStyles");
            xml.attribute("count", "1");
            xml.emptyElement("cellStyle");
            xml.attribute("name", "Normal");
            xml.attribute("xfId", "0");
            xml.attribute("builtinId", "0");
            xml.endElement();

            xml.endElement(); // styleSheet
            xml.endDocument();
        }
    }

    private static void writeFont(XmlWriter xml, Font font) throws IOException {
        xml.startElement("font");

        if (font.bold()) {
            xml.emptyElement("b");
        }
        if (font.italic()) {
            xml.emptyElement("i");
        }
        if (font.underline()) {
            xml.emptyElement("u");
        }
        if (font.strikethrough()) {
            xml.emptyElement("strike");
        }

        xml.emptyElement("sz");
        xml.attribute("val", String.valueOf(font.size()));

        if (font.color() != 0 && font.color() != 0xFF000000) {
            xml.emptyElement("color");
            xml.attribute("rgb", toArgbString(font.color()));
        } else {
            xml.emptyElement("color");
            xml.attribute("theme", "1");
        }

        xml.emptyElement("name");
        xml.attribute("val", font.name());

        xml.emptyElement("family");
        xml.attribute("val", "2");

        xml.emptyElement("scheme");
        xml.attribute("val", "minor");

        xml.endElement();
    }

    private static void writeBorder(XmlWriter xml, Border border) throws IOException {
        xml.startElement("border");

        writeBorderSide(xml, "left", border.left());
        writeBorderSide(xml, "right", border.right());
        writeBorderSide(xml, "top", border.top());
        writeBorderSide(xml, "bottom", border.bottom());
        xml.emptyElement("diagonal");

        xml.endElement();
    }

    private static void writeBorderSide(XmlWriter xml, String name, Border.BorderSide side) throws IOException {
        if (side == null || side.style() == BorderStyle.NONE) {
            xml.emptyElement(name);
        } else {
            xml.startElement(name);
            xml.attribute("style", borderStyleName(side.style()));
            if (side.color() != 0) {
                xml.emptyElement("color");
                xml.attribute("rgb", toArgbString(side.color()));
            }
            xml.endElement();
        }
    }

    private static String borderStyleName(BorderStyle style) {
        return switch (style) {
            case NONE -> "none";
            case THIN -> "thin";
            case MEDIUM -> "medium";
            case THICK -> "thick";
            case DOUBLE -> "double";
            case DASHED -> "dashed";
            case DOTTED -> "dotted";
        };
    }

    private static String toArgbString(int argb) {
        return String.format("%08X", argb);
    }

    private static <T> int indexOfOrAdd(List<T> list, T item) {
        for (int i = 0; i < list.size(); i++) {
            if (Objects.equals(list.get(i), item)) {
                return i;
            }
        }
        list.add(item);
        return list.size() - 1;
    }

    private static int indexOfFillOrAdd(List<Integer> fills, int color) {
        if (color == 0) {
            return 0; // No fill
        }
        for (int i = 2; i < fills.size(); i++) {
            if (fills.get(i) == color) {
                return i;
            }
        }
        fills.add(color);
        return fills.size() - 1;
    }
}
