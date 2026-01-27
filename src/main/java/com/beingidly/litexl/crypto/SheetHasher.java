package com.beingidly.litexl.crypto;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.security.SecureRandom;
import java.util.Base64;

/**
 * Password hashing for sheet protection.
 * Implements the modern SHA-512 based algorithm.
 *
 * <p>Usage:</p>
 * <pre>{@code
 * SheetHasher hasher = new SheetHasher();
 * SheetProtectionInfo info = hasher.hash("password");
 * boolean valid = hasher.verify("password", info);
 * }</pre>
 */
public final class SheetHasher {

    private static final int DEFAULT_SPIN_COUNT = 100000;

    private final SecureRandom random = new SecureRandom();
    private final int spinCount;

    /**
     * Creates a new sheet hasher with the default spin count (100,000).
     */
    public SheetHasher() {
        this(DEFAULT_SPIN_COUNT);
    }

    /**
     * Creates a new sheet hasher with a custom spin count.
     *
     * @param spinCount the number of hash iterations
     */
    public SheetHasher(int spinCount) {
        this.spinCount = spinCount;
    }

    /**
     * Creates a password hash for sheet protection.
     *
     * @param password the password to hash
     * @return sheet protection info with salt, hash, and algorithm parameters
     */
    public SheetProtectionInfo hash(String password) {
        byte[] salt = new byte[16];
        random.nextBytes(salt);

        byte[] hash = computeHash(password, salt, spinCount);

        return new SheetProtectionInfo(
            Base64.getEncoder().encodeToString(salt),
            Base64.getEncoder().encodeToString(hash),
            "SHA-512",
            spinCount
        );
    }

    /**
     * Verifies a password against stored protection info.
     *
     * @param password the password to verify
     * @param info the stored protection info
     * @return true if the password matches
     */
    public boolean verify(String password, SheetProtectionInfo info) {
        byte[] salt = Base64.getDecoder().decode(info.saltValue());
        byte[] expectedHash = Base64.getDecoder().decode(info.hashValue());

        byte[] computedHash = computeHash(password, salt, info.spinCount());

        return MessageDigest.isEqual(expectedHash, computedHash);
    }

    private byte[] computeHash(String password, byte[] salt, int iterations) {
        try {
            MessageDigest sha512 = MessageDigest.getInstance("SHA-512");

            // Initial hash: H0 = SHA512(salt + password)
            byte[] passwordBytes = password.getBytes(StandardCharsets.UTF_16LE);
            sha512.update(salt);
            sha512.update(passwordBytes);
            byte[] hash = sha512.digest();

            // Iterate: Hn = SHA512(iterator + Hn-1)
            byte[] iterBuf = new byte[4];
            for (int i = 0; i < iterations; i++) {
                iterBuf[0] = (byte) i;
                iterBuf[1] = (byte) (i >> 8);
                iterBuf[2] = (byte) (i >> 16);
                iterBuf[3] = (byte) (i >> 24);
                sha512.update(iterBuf);
                sha512.update(hash);
                hash = sha512.digest();
            }

            return hash;
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException("SHA-512 not available", e);
        }
    }

    /**
     * Sheet protection information.
     */
    public record SheetProtectionInfo(
        String saltValue,
        String hashValue,
        String algorithmName,
        int spinCount
    ) {}
}
