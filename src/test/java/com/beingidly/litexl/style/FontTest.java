package com.beingidly.litexl.style;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class FontTest {

    @Test
    void defaultFont() {
        Font font = Font.DEFAULT;

        assertEquals("Calibri", font.name());
        assertEquals(11.0, font.size());
        assertFalse(font.bold());
        assertFalse(font.italic());
        assertFalse(font.underline());
        assertFalse(font.strikethrough());
        assertEquals(0xFF000000, font.color());
    }

    @Test
    void fontOf() {
        Font font = Font.of("Arial", 14);

        assertEquals("Arial", font.name());
        assertEquals(14.0, font.size());
        assertEquals(0xFF000000, font.color());
        assertFalse(font.bold());
    }

    @Test
    void withBold() {
        Font font = Font.DEFAULT.withBold(true);

        assertTrue(font.bold());
        assertEquals("Calibri", font.name());
        assertEquals(11.0, font.size());
    }

    @Test
    void withItalic() {
        Font font = Font.DEFAULT.withItalic(true);

        assertTrue(font.italic());
    }

    @Test
    void withColor() {
        Font font = Font.DEFAULT.withColor(0xFFFF0000);

        assertEquals(0xFFFF0000, font.color());
    }

    @Test
    void fontChaining() {
        Font font = Font.DEFAULT
            .withBold(true)
            .withItalic(true)
            .withColor(0xFF0000FF);

        assertTrue(font.bold());
        assertTrue(font.italic());
        assertEquals(0xFF0000FF, font.color());
    }

    @Test
    void fontEquality() {
        Font font1 = Font.of("Arial", 12);
        Font font2 = Font.of("Arial", 12);
        Font font3 = Font.of("Calibri", 12);

        assertEquals(font1, font2);
        assertEquals(font1.hashCode(), font2.hashCode());
        assertNotEquals(font1, font3);
    }

    @Test
    void customFont() {
        Font font = new Font("Times New Roman", 16, 0xFF0000FF, true, true, true, true);

        assertEquals("Times New Roman", font.name());
        assertEquals(16.0, font.size());
        assertEquals(0xFF0000FF, font.color());
        assertTrue(font.bold());
        assertTrue(font.italic());
        assertTrue(font.underline());
        assertTrue(font.strikethrough());
    }
}
