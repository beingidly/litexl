package com.beingidly.litexl.crypto;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class EncryptionOptionsTest {

    @Test
    void aes256CreatesWithDefaultSpinCount() {
        EncryptionOptions options = EncryptionOptions.aes256("password");

        assertEquals(EncryptionOptions.Algorithm.AES_256, options.algorithm());
        assertEquals("password", options.password());
        assertEquals(100000, options.spinCount());
    }

    @Test
    void aes128CreatesWithDefaultSpinCount() {
        EncryptionOptions options = EncryptionOptions.aes128("secret");

        assertEquals(EncryptionOptions.Algorithm.AES_128, options.algorithm());
        assertEquals("secret", options.password());
        assertEquals(100000, options.spinCount());
    }

    @Test
    void aes256WithCustomSpinCount() {
        EncryptionOptions options = EncryptionOptions.aes256("password", 50000);

        assertEquals(EncryptionOptions.Algorithm.AES_256, options.algorithm());
        assertEquals("password", options.password());
        assertEquals(50000, options.spinCount());
    }

    @Test
    void constructorWithAllParameters() {
        EncryptionOptions options = new EncryptionOptions(
            EncryptionOptions.Algorithm.AES_128, "myPassword", 200000);

        assertEquals(EncryptionOptions.Algorithm.AES_128, options.algorithm());
        assertEquals("myPassword", options.password());
        assertEquals(200000, options.spinCount());
    }

    @Test
    void nullPasswordThrows() {
        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class, () ->
            EncryptionOptions.aes256(null));

        assertEquals("Password cannot be null or empty", ex.getMessage());
    }

    @Test
    void emptyPasswordThrows() {
        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class, () ->
            EncryptionOptions.aes256(""));

        assertEquals("Password cannot be null or empty", ex.getMessage());
    }

    @Test
    void zeroSpinCountThrows() {
        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class, () ->
            new EncryptionOptions(EncryptionOptions.Algorithm.AES_256, "password", 0));

        assertEquals("Spin count must be positive", ex.getMessage());
    }

    @Test
    void negativeSpinCountThrows() {
        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class, () ->
            new EncryptionOptions(EncryptionOptions.Algorithm.AES_256, "password", -1));

        assertEquals("Spin count must be positive", ex.getMessage());
    }

    @Test
    void spinCountOfOneIsValid() {
        EncryptionOptions options = new EncryptionOptions(
            EncryptionOptions.Algorithm.AES_256, "password", 1);

        assertEquals(1, options.spinCount());
    }

    @Test
    void algorithmEnumValues() {
        EncryptionOptions.Algorithm[] values = EncryptionOptions.Algorithm.values();

        assertEquals(2, values.length);
        assertEquals(EncryptionOptions.Algorithm.AES_128, EncryptionOptions.Algorithm.valueOf("AES_128"));
        assertEquals(EncryptionOptions.Algorithm.AES_256, EncryptionOptions.Algorithm.valueOf("AES_256"));
    }

    @Test
    void recordEquality() {
        EncryptionOptions options1 = EncryptionOptions.aes256("password");
        EncryptionOptions options2 = EncryptionOptions.aes256("password");
        EncryptionOptions options3 = EncryptionOptions.aes256("different");

        assertEquals(options1, options2);
        assertNotEquals(options1, options3);
    }

    @Test
    void recordHashCode() {
        EncryptionOptions options1 = EncryptionOptions.aes256("password");
        EncryptionOptions options2 = EncryptionOptions.aes256("password");

        assertEquals(options1.hashCode(), options2.hashCode());
    }

    @Test
    void recordToString() {
        EncryptionOptions options = EncryptionOptions.aes256("secret");
        String str = options.toString();

        assertTrue(str.contains("AES_256"));
        assertTrue(str.contains("secret"));
        assertTrue(str.contains("100000"));
    }

    @Test
    void differentAlgorithmsNotEqual() {
        EncryptionOptions aes128 = EncryptionOptions.aes128("password");
        EncryptionOptions aes256 = EncryptionOptions.aes256("password");

        assertNotEquals(aes128, aes256);
    }

    @Test
    void differentSpinCountsNotEqual() {
        EncryptionOptions options1 = EncryptionOptions.aes256("password", 100000);
        EncryptionOptions options2 = EncryptionOptions.aes256("password", 50000);

        assertNotEquals(options1, options2);
    }
}
