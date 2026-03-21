package com.beingidly.litexl.crypto;

/**
 * Options for encrypting a workbook.
 *
 * @param algorithm the AES algorithm variant to use
 * @param password the encryption password
 * @param spinCount the number of hash iterations for key derivation
 */
public record EncryptionOptions(
    Algorithm algorithm,
    String password,
    int spinCount
) {
    /** AES algorithm variants. */
    public enum Algorithm {
        /** AES with 128-bit key. */
        AES_128,
        /** AES with 256-bit key. */
        AES_256
    }

    /**
     * Validates that password is not empty and spin count is positive.
     */
    public EncryptionOptions {
        if (password == null || password.isEmpty()) {
            throw new IllegalArgumentException("Password cannot be null or empty");
        }
        if (spinCount < 1) {
            throw new IllegalArgumentException("Spin count must be positive");
        }
    }

    /**
     * Creates AES-256 encryption options with default spin count.
     *
     * @param password the encryption password
     * @return new AES-256 encryption options
     */
    public static EncryptionOptions aes256(String password) {
        return new EncryptionOptions(Algorithm.AES_256, password, 100000);
    }

    /**
     * Creates AES-128 encryption options with default spin count.
     *
     * @param password the encryption password
     * @return new AES-128 encryption options
     */
    public static EncryptionOptions aes128(String password) {
        return new EncryptionOptions(Algorithm.AES_128, password, 100000);
    }

    /**
     * Creates AES-256 encryption options with custom spin count.
     *
     * @param password the encryption password
     * @param spinCount the number of hash iterations
     * @return new AES-256 encryption options
     */
    public static EncryptionOptions aes256(String password, int spinCount) {
        return new EncryptionOptions(Algorithm.AES_256, password, spinCount);
    }
}
