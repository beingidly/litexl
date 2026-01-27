package com.beingidly.litexl.crypto;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class SheetHasherTest {

    @Test
    void hash() {
        SheetHasher hasher = new SheetHasher();
        SheetHasher.SheetProtectionInfo info = hasher.hash("password123");

        assertNotNull(info.saltValue());
        assertNotNull(info.hashValue());
        assertEquals("SHA-512", info.algorithmName());
        assertEquals(100000, info.spinCount());
    }

    @Test
    void hashWithCustomSpinCount() {
        SheetHasher hasher = new SheetHasher(50000);
        SheetHasher.SheetProtectionInfo info = hasher.hash("password");

        assertNotNull(info.saltValue());
        assertNotNull(info.hashValue());
        assertEquals(50000, info.spinCount());
    }

    @Test
    void verify() {
        SheetHasher hasher = new SheetHasher();
        SheetHasher.SheetProtectionInfo info = hasher.hash("mypassword");

        assertTrue(hasher.verify("mypassword", info));
        assertFalse(hasher.verify("wrongpassword", info));
    }

    @Test
    void differentSaltsProduceDifferentHashes() {
        SheetHasher hasher = new SheetHasher();
        SheetHasher.SheetProtectionInfo info1 = hasher.hash("password");
        SheetHasher.SheetProtectionInfo info2 = hasher.hash("password");

        // Salt should be different (random)
        assertNotEquals(info1.saltValue(), info2.saltValue());
        // Hash should be different due to different salt
        assertNotEquals(info1.hashValue(), info2.hashValue());
    }
}
