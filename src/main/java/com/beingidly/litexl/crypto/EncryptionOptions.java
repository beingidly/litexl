package com.beingidly.litexl.crypto;

/**
 * Options for encrypting a workbook.
 */
public record EncryptionOptions(
    Algorithm algorithm,
    String password,
    int spinCount
) {
    public enum Algorithm {
        AES_128,
        AES_256
    }

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
     */
    public static EncryptionOptions aes256(String password) {
        return new EncryptionOptions(Algorithm.AES_256, password, 100000);
    }

    /**
     * Creates AES-128 encryption options with default spin count.
     */
    public static EncryptionOptions aes128(String password) {
        return new EncryptionOptions(Algorithm.AES_128, password, 100000);
    }

    /**
     * Creates AES-256 encryption options with custom spin count.
     */
    public static EncryptionOptions aes256(String password, int spinCount) {
        return new EncryptionOptions(Algorithm.AES_256, password, spinCount);
    }
}
