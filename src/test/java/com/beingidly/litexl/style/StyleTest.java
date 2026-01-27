package com.beingidly.litexl.style;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class StyleTest {

    @Test
    void defaultStyle() {
        Style style = Style.DEFAULT;

        assertEquals(Font.DEFAULT, style.font());
        assertEquals(Alignment.DEFAULT, style.alignment());
        assertEquals(0, style.fillColor());
        assertEquals(Border.NONE, style.border());
        assertNull(style.numberFormat());
        assertTrue(style.locked());
        assertFalse(style.wrapText());
    }

    @Test
    void builderWithFont() {
        Font font = Font.of("Arial", 14).withBold(true);

        Style style = Style.builder()
            .font(font)
            .build();

        assertEquals(font, style.font());
    }

    @Test
    void builderWithFontShorthand() {
        Style style = Style.builder()
            .font("Arial", 14)
            .build();

        assertEquals("Arial", style.font().name());
        assertEquals(14.0, style.font().size());
    }

    @Test
    void builderWithAlignment() {
        Alignment align = Alignment.of(HAlign.CENTER, VAlign.MIDDLE);

        Style style = Style.builder()
            .alignment(align)
            .build();

        assertEquals(align, style.alignment());
    }

    @Test
    void builderWithFillColor() {
        Style style = Style.builder()
            .fill(0xFFFF0000)
            .build();

        assertEquals(0xFFFF0000, style.fillColor());
    }

    @Test
    void builderWithBorder() {
        Border border = Border.all(BorderStyle.THIN, 0xFF000000);

        Style style = Style.builder()
            .border(border)
            .build();

        assertEquals(border, style.border());
    }

    @Test
    void builderWithBorderShorthand() {
        Style style = Style.builder()
            .border(BorderStyle.THIN, 0xFF000000)
            .build();

        assertEquals(BorderStyle.THIN, style.border().top().style());
    }

    @Test
    void builderWithNumberFormat() {
        Style style = Style.builder()
            .format("#,##0.00")
            .build();

        assertEquals("#,##0.00", style.numberFormat());
    }

    @Test
    void builderWithLocked() {
        Style style = Style.builder()
            .locked(false)
            .build();

        assertFalse(style.locked());
    }

    @Test
    void builderWithWrapText() {
        Style style = Style.builder()
            .wrap(true)
            .build();

        assertTrue(style.wrapText());
    }

    @Test
    void builderWithBold() {
        Style style = Style.builder()
            .bold(true)
            .build();

        assertTrue(style.font().bold());
    }

    @Test
    void builderWithItalic() {
        Style style = Style.builder()
            .italic(true)
            .build();

        assertTrue(style.font().italic());
    }

    @Test
    void builderWithUnderline() {
        Style style = Style.builder()
            .underline(true)
            .build();

        assertTrue(style.font().underline());
    }

    @Test
    void builderWithFontColor() {
        Style style = Style.builder()
            .color(0xFFFF0000)
            .build();

        assertEquals(0xFFFF0000, style.font().color());
    }

    @Test
    void builderWithAlign() {
        Style style = Style.builder()
            .align(HAlign.RIGHT, VAlign.TOP)
            .build();

        assertEquals(HAlign.RIGHT, style.alignment().horizontal());
        assertEquals(VAlign.TOP, style.alignment().vertical());
    }

    @Test
    void builderWithBorderSides() {
        Style style = Style.builder()
            .borderTop(BorderStyle.THIN, 0xFF000000)
            .borderBottom(BorderStyle.MEDIUM, 0xFF000000)
            .borderLeft(BorderStyle.THICK, 0xFF000000)
            .borderRight(BorderStyle.DASHED, 0xFF000000)
            .build();

        assertEquals(BorderStyle.THIN, style.border().top().style());
        assertEquals(BorderStyle.MEDIUM, style.border().bottom().style());
        assertEquals(BorderStyle.THICK, style.border().left().style());
        assertEquals(BorderStyle.DASHED, style.border().right().style());
    }

    @Test
    void fullStyleBuilder() {
        Style style = Style.builder()
            .font(Font.of("Arial", 14).withBold(true))
            .alignment(Alignment.of(HAlign.CENTER, VAlign.TOP))
            .fill(0xFF00FF00)
            .border(Border.all(BorderStyle.MEDIUM, 0xFF0000FF))
            .format("0.00%")
            .locked(false)
            .wrap(true)
            .build();

        assertTrue(style.font().bold());
        assertEquals(HAlign.CENTER, style.alignment().horizontal());
        assertEquals(VAlign.TOP, style.alignment().vertical());
        assertEquals(0xFF00FF00, style.fillColor());
        assertEquals(BorderStyle.MEDIUM, style.border().top().style());
        assertEquals("0.00%", style.numberFormat());
        assertFalse(style.locked());
        assertTrue(style.wrapText());
    }

    @Test
    void styleEquality() {
        Style style1 = Style.builder()
            .fill(0xFFFF0000)
            .build();

        Style style2 = Style.builder()
            .fill(0xFFFF0000)
            .build();

        Style style3 = Style.builder()
            .fill(0xFF00FF00)
            .build();

        assertEquals(style1, style2);
        assertEquals(style1.hashCode(), style2.hashCode());
        assertNotEquals(style1, style3);
    }
}
