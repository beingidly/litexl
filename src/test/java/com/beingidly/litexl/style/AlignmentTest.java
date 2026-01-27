package com.beingidly.litexl.style;

import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class AlignmentTest {

    @Test
    void default_hasGeneralAndBottom() {
        assertEquals(HAlign.GENERAL, Alignment.DEFAULT.horizontal());
        assertEquals(VAlign.BOTTOM, Alignment.DEFAULT.vertical());
    }

    @Test
    void of_createsAlignment() {
        Alignment align = Alignment.of(HAlign.LEFT, VAlign.TOP);
        assertEquals(HAlign.LEFT, align.horizontal());
        assertEquals(VAlign.TOP, align.vertical());
    }

    @Test
    void center_hasCenterMiddle() {
        Alignment center = Alignment.center();
        assertEquals(HAlign.CENTER, center.horizontal());
        assertEquals(VAlign.MIDDLE, center.vertical());
    }

    @Test
    void leftTop_hasLeftTop() {
        Alignment leftTop = Alignment.leftTop();
        assertEquals(HAlign.LEFT, leftTop.horizontal());
        assertEquals(VAlign.TOP, leftTop.vertical());
    }

    @Test
    void record_equality() {
        Alignment a1 = Alignment.of(HAlign.RIGHT, VAlign.MIDDLE);
        Alignment a2 = Alignment.of(HAlign.RIGHT, VAlign.MIDDLE);
        assertEquals(a1, a2);
        assertEquals(a1.hashCode(), a2.hashCode());
    }
}
