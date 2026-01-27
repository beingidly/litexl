package com.beingidly.litexl.style;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class BorderTest {

    @Test
    void noneBorder() {
        Border border = Border.NONE;

        assertEquals(BorderStyle.NONE, border.top().style());
        assertEquals(BorderStyle.NONE, border.bottom().style());
        assertEquals(BorderStyle.NONE, border.left().style());
        assertEquals(BorderStyle.NONE, border.right().style());
    }

    @Test
    void allBorder() {
        Border border = Border.all(BorderStyle.THIN, 0xFF000000);

        assertEquals(BorderStyle.THIN, border.top().style());
        assertEquals(BorderStyle.THIN, border.bottom().style());
        assertEquals(BorderStyle.THIN, border.left().style());
        assertEquals(BorderStyle.THIN, border.right().style());

        assertEquals(0xFF000000, border.top().color());
        assertEquals(0xFF000000, border.bottom().color());
        assertEquals(0xFF000000, border.left().color());
        assertEquals(0xFF000000, border.right().color());
    }

    @Test
    void thinBlackBorder() {
        Border border = Border.thinBlack();

        assertEquals(BorderStyle.THIN, border.top().style());
        assertEquals(BorderStyle.THIN, border.bottom().style());
        assertEquals(BorderStyle.THIN, border.left().style());
        assertEquals(BorderStyle.THIN, border.right().style());

        assertEquals(0xFF000000, border.top().color());
    }

    @Test
    void borderSideThin() {
        Border.BorderSide side = Border.BorderSide.thin(0xFFFF0000);

        assertEquals(BorderStyle.THIN, side.style());
        assertEquals(0xFFFF0000, side.color());
    }

    @Test
    void borderSideMedium() {
        Border.BorderSide side = Border.BorderSide.medium(0xFF00FF00);

        assertEquals(BorderStyle.MEDIUM, side.style());
        assertEquals(0xFF00FF00, side.color());
    }

    @Test
    void borderSideThick() {
        Border.BorderSide side = Border.BorderSide.thick(0xFF0000FF);

        assertEquals(BorderStyle.THICK, side.style());
        assertEquals(0xFF0000FF, side.color());
    }

    @Test
    void borderSideNone() {
        Border.BorderSide side = Border.BorderSide.NONE;

        assertEquals(BorderStyle.NONE, side.style());
        assertEquals(0, side.color());
    }

    @Test
    void customBorder() {
        Border border = new Border(
            Border.BorderSide.thin(0xFF000000),
            Border.BorderSide.medium(0xFF000000),
            Border.BorderSide.thick(0xFF000000),
            Border.BorderSide.NONE
        );

        assertEquals(BorderStyle.THIN, border.left().style());
        assertEquals(BorderStyle.MEDIUM, border.right().style());
        assertEquals(BorderStyle.THICK, border.top().style());
        assertEquals(BorderStyle.NONE, border.bottom().style());
    }

    @Test
    void borderStylesAvailable() {
        BorderStyle[] styles = BorderStyle.values();

        assertTrue(styles.length >= 7); // NONE, THIN, MEDIUM, THICK, DOUBLE, DASHED, DOTTED
    }

    @Test
    void borderEquality() {
        Border border1 = Border.all(BorderStyle.THIN, 0xFF000000);
        Border border2 = Border.all(BorderStyle.THIN, 0xFF000000);
        Border border3 = Border.all(BorderStyle.MEDIUM, 0xFF000000);

        assertEquals(border1, border2);
        assertEquals(border1.hashCode(), border2.hashCode());
        assertNotEquals(border1, border3);
    }
}
