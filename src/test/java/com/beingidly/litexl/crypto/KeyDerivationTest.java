package com.beingidly.litexl.crypto;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class KeyDerivationTest {

    @Test
    void deriveKey128Bit() {
        byte[] salt = new byte[16];
        salt[0] = 1;
        salt[1] = 2;
        salt[2] = 3;

        byte[] key = KeyDerivation.deriveKey(
            "password",
            salt,
            100,
            128,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        assertEquals(16, key.length); // 128 bits = 16 bytes
    }

    @Test
    void deriveKey256Bit() {
        byte[] salt = new byte[16];
        salt[0] = 1;
        salt[1] = 2;
        salt[2] = 3;

        byte[] key = KeyDerivation.deriveKey(
            "password",
            salt,
            100,
            256,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        assertEquals(32, key.length); // 256 bits = 32 bytes
    }

    @Test
    void sameInputProducesSameKey() {
        byte[] salt = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16};

        byte[] key1 = KeyDerivation.deriveKey(
            "password",
            salt,
            1000,
            256,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        byte[] key2 = KeyDerivation.deriveKey(
            "password",
            salt,
            1000,
            256,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        assertArrayEquals(key1, key2);
    }

    @Test
    void differentPasswordsProduceDifferentKeys() {
        byte[] salt = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16};

        byte[] key1 = KeyDerivation.deriveKey(
            "password1",
            salt,
            1000,
            256,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        byte[] key2 = KeyDerivation.deriveKey(
            "password2",
            salt,
            1000,
            256,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        assertFalse(java.util.Arrays.equals(key1, key2));
    }

    @Test
    void differentSaltsProduceDifferentKeys() {
        byte[] salt1 = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16};
        byte[] salt2 = {16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1};

        byte[] key1 = KeyDerivation.deriveKey(
            "password",
            salt1,
            1000,
            256,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        byte[] key2 = KeyDerivation.deriveKey(
            "password",
            salt2,
            1000,
            256,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        assertFalse(java.util.Arrays.equals(key1, key2));
    }

    @Test
    void differentBlockKeysProduceDifferentKeys() {
        byte[] salt = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16};

        byte[] key1 = KeyDerivation.deriveKey(
            "password",
            salt,
            1000,
            256,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        byte[] key2 = KeyDerivation.deriveKey(
            "password",
            salt,
            1000,
            256,
            KeyDerivation.BLOCK_KEY_VERIFIER_INPUT
        );

        assertFalse(java.util.Arrays.equals(key1, key2));
    }

    @Test
    void blockKeysHaveCorrectLength() {
        assertEquals(8, KeyDerivation.BLOCK_KEY_VERIFIER_INPUT.length);
        assertEquals(8, KeyDerivation.BLOCK_KEY_VERIFIER_VALUE.length);
        assertEquals(8, KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY.length);
        assertEquals(8, KeyDerivation.BLOCK_KEY_DATA_INTEGRITY_HMAC_KEY.length);
        assertEquals(8, KeyDerivation.BLOCK_KEY_DATA_INTEGRITY_HMAC_VALUE.length);
    }

    @Test
    void batchDerivationMatchesSingleDerivation() {
        String password = "testPassword123";
        byte[] salt = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16};
        int spinCount = 1000;
        int keyBits = 256;

        // Derive keys using single derivation (3 separate calls)
        byte[] key1Single = KeyDerivation.deriveKey(
            password, salt, spinCount, keyBits,
            KeyDerivation.BLOCK_KEY_VERIFIER_INPUT
        );
        byte[] key2Single = KeyDerivation.deriveKey(
            password, salt, spinCount, keyBits,
            KeyDerivation.BLOCK_KEY_VERIFIER_VALUE
        );
        byte[] key3Single = KeyDerivation.deriveKey(
            password, salt, spinCount, keyBits,
            KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        // Derive keys using batch derivation (1 intermediate + 3 final)
        byte[] intermediate = KeyDerivation.deriveIntermediateHash(password, salt, spinCount);
        byte[] key1Batch = KeyDerivation.deriveKeyFromIntermediate(
            intermediate, keyBits, KeyDerivation.BLOCK_KEY_VERIFIER_INPUT
        );
        byte[] key2Batch = KeyDerivation.deriveKeyFromIntermediate(
            intermediate, keyBits, KeyDerivation.BLOCK_KEY_VERIFIER_VALUE
        );
        byte[] key3Batch = KeyDerivation.deriveKeyFromIntermediate(
            intermediate, keyBits, KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        // Verify batch results match single derivation results
        assertArrayEquals(key1Single, key1Batch, "VERIFIER_INPUT key should match");
        assertArrayEquals(key2Single, key2Batch, "VERIFIER_VALUE key should match");
        assertArrayEquals(key3Single, key3Batch, "ENCRYPTED_KEY key should match");
    }

    @Test
    void deriveKeyProducesDeterministicOutput() {
        String password = "deterministicTest";
        byte[] salt = {10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 13, 14, 15, 16};
        int spinCount = 500;
        int keyBits = 128;

        // Call deriveKey twice with identical inputs
        byte[] key1 = KeyDerivation.deriveKey(
            password, salt, spinCount, keyBits,
            KeyDerivation.BLOCK_KEY_DATA_INTEGRITY_HMAC_KEY
        );
        byte[] key2 = KeyDerivation.deriveKey(
            password, salt, spinCount, keyBits,
            KeyDerivation.BLOCK_KEY_DATA_INTEGRITY_HMAC_KEY
        );

        assertArrayEquals(key1, key2, "Same inputs should produce identical keys");

        // Also verify intermediate hash is deterministic
        byte[] intermediate1 = KeyDerivation.deriveIntermediateHash(password, salt, spinCount);
        byte[] intermediate2 = KeyDerivation.deriveIntermediateHash(password, salt, spinCount);

        assertArrayEquals(intermediate1, intermediate2, "Same inputs should produce identical intermediate hashes");
    }

    @Test
    void intermediateHashHasCorrectLength() {
        String password = "test";
        byte[] salt = new byte[16];
        int spinCount = 100;

        byte[] intermediate = KeyDerivation.deriveIntermediateHash(password, salt, spinCount);

        assertEquals(64, intermediate.length, "Intermediate hash should be 64 bytes (SHA-512 output)");
    }

    @Test
    void deriveKeyFromIntermediateProducesCorrectKeyLength() {
        byte[] intermediate = new byte[64]; // Simulated intermediate hash
        intermediate[0] = 1;
        intermediate[1] = 2;

        byte[] key128 = KeyDerivation.deriveKeyFromIntermediate(
            intermediate, 128, KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );
        byte[] key256 = KeyDerivation.deriveKeyFromIntermediate(
            intermediate, 256, KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        assertEquals(16, key128.length, "128-bit key should be 16 bytes");
        assertEquals(32, key256.length, "256-bit key should be 32 bytes");
    }
}
