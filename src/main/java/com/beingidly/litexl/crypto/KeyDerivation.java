package com.beingidly.litexl.crypto;

import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.Arrays;

/**
 * Key derivation functions for ECMA-376 encryption.
 */
public final class KeyDerivation {

    private KeyDerivation() {}

    /**
     * Derives a key using the ECMA-376 Agile Encryption algorithm.
     *
     * @param password The password
     * @param salt The salt
     * @param spinCount Number of hash iterations
     * @param keyBits Key size in bits (128 or 256)
     * @param blockKey Block key for deriving specific keys
     * @return The derived key
     */
    public static byte[] deriveKey(String password, byte[] salt, int spinCount, int keyBits, byte[] blockKey) {
        byte[] intermediateHash = deriveIntermediateHash(password, salt, spinCount);
        return deriveKeyFromIntermediate(intermediateHash, keyBits, blockKey);
    }

    /**
     * Derives the intermediate hash after spinCount iterations.
     * Use with deriveKeyFromIntermediate() for multiple block keys.
     *
     * <p>This method performs the expensive iteration loop once. When deriving
     * multiple keys with the same password/salt/spinCount but different block keys,
     * call this method once and then use deriveKeyFromIntermediate() for each block key.
     *
     * @param password The password
     * @param salt The salt
     * @param spinCount Number of hash iterations
     * @return The 64-byte intermediate hash
     */
    public static byte[] deriveIntermediateHash(String password, byte[] salt, int spinCount) {
        try {
            MessageDigest sha512 = MessageDigest.getInstance("SHA-512");

            // Initial hash: H0 = SHA512(salt + password)
            byte[] passwordBytes = password.getBytes(StandardCharsets.UTF_16LE);
            sha512.update(salt);
            sha512.update(passwordBytes);
            byte[] hash = sha512.digest();

            // Iterate: Hn = SHA512(iterator + Hn-1)
            ByteBuffer iterBuf = ByteBuffer.allocate(4).order(ByteOrder.LITTLE_ENDIAN);
            for (int i = 0; i < spinCount; i++) {
                iterBuf.clear();
                iterBuf.putInt(i);
                sha512.update(iterBuf.array());
                sha512.update(hash);
                hash = sha512.digest();
            }

            return hash;

        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException("SHA-512 not available", e);
        }
    }

    /**
     * Derives a key from pre-computed intermediate hash.
     * Much faster than full deriveKey() when deriving multiple keys.
     *
     * @param intermediateHash The 64-byte intermediate hash from deriveIntermediateHash()
     * @param keyBits Key size in bits (128 or 256)
     * @param blockKey Block key for deriving specific keys
     * @return The derived key
     */
    public static byte[] deriveKeyFromIntermediate(byte[] intermediateHash, int keyBits, byte[] blockKey) {
        try {
            MessageDigest sha512 = MessageDigest.getInstance("SHA-512");

            // Derive final key: Hfinal = SHA512(Hn + blockKey)
            sha512.update(intermediateHash);
            sha512.update(blockKey);
            byte[] derivedKey = sha512.digest();

            // Truncate to key size
            int keyBytes = keyBits / 8;
            return Arrays.copyOf(derivedKey, keyBytes);

        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException("SHA-512 not available", e);
        }
    }

    /**
     * Block keys used in ECMA-376 Agile Encryption.
     */
    public static final byte[] BLOCK_KEY_VERIFIER_INPUT = {
        (byte) 0xfe, (byte) 0xa7, (byte) 0xd2, (byte) 0x76,
        (byte) 0x3b, (byte) 0x4b, (byte) 0x9e, (byte) 0x79
    };

    public static final byte[] BLOCK_KEY_VERIFIER_VALUE = {
        (byte) 0xd7, (byte) 0xaa, (byte) 0x0f, (byte) 0x6d,
        (byte) 0x30, (byte) 0x61, (byte) 0x34, (byte) 0x4e
    };

    public static final byte[] BLOCK_KEY_ENCRYPTED_KEY = {
        (byte) 0x14, (byte) 0x6e, (byte) 0x0b, (byte) 0xe7,
        (byte) 0xab, (byte) 0xac, (byte) 0xd0, (byte) 0xd6
    };

    public static final byte[] BLOCK_KEY_DATA_INTEGRITY_HMAC_KEY = {
        (byte) 0x5f, (byte) 0xb2, (byte) 0xad, (byte) 0x01,
        (byte) 0x0c, (byte) 0xb9, (byte) 0xe1, (byte) 0xf6
    };

    public static final byte[] BLOCK_KEY_DATA_INTEGRITY_HMAC_VALUE = {
        (byte) 0xa0, (byte) 0x67, (byte) 0x7f, (byte) 0x02,
        (byte) 0xb2, (byte) 0x2c, (byte) 0x84, (byte) 0x33
    };
}
