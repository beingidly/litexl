package com.beingidly.litexl.crypto;

import javax.crypto.Cipher;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import java.nio.ByteBuffer;
import java.security.GeneralSecurityException;

/**
 * AES encryption/decryption with reusable Cipher and SecretKeySpec.
 *
 * <p>Instance-based design eliminates per-call allocation of SecretKeySpec,
 * which is critical for segment-by-segment encryption/decryption.</p>
 *
 * <p>Usage:</p>
 * <pre>{@code
 * AesCipher cipher = new AesCipher(key);
 * cipher.encrypt(input, output, iv);
 * }</pre>
 */
public final class AesCipher {

    private final Cipher cipher;
    private final SecretKeySpec keySpec;

    /**
     * Creates a new AES cipher with the given key.
     *
     * @param key AES key (16, 24, or 32 bytes)
     */
    public AesCipher(byte[] key) {
        try {
            this.cipher = Cipher.getInstance("AES/CBC/NoPadding");
        } catch (GeneralSecurityException e) {
            throw new RuntimeException("AES cipher not available", e);
        }
        this.keySpec = new SecretKeySpec(key, "AES");
    }

    /**
     * Encrypts data using AES-CBC with zero padding to block size.
     *
     * @param data the data to encrypt
     * @param iv initialization vector (16 bytes)
     * @return encrypted data
     */
    public byte[] encrypt(byte[] data, byte[] iv) throws GeneralSecurityException {
        return encrypt(data, 0, data.length, iv);
    }

    /**
     * Encrypts a portion of data using AES-CBC with zero padding to block size.
     *
     * @param data the data to encrypt
     * @param offset start offset in data
     * @param length number of bytes to encrypt
     * @param iv initialization vector (16 bytes)
     * @return encrypted data
     */
    public byte[] encrypt(byte[] data, int offset, int length, byte[] iv) throws GeneralSecurityException {
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.ENCRYPT_MODE, keySpec, ivSpec);

        // Pad to block size
        int blockSize = 16;
        int paddedLength = ((length + blockSize - 1) / blockSize) * blockSize;
        byte[] padded = new byte[paddedLength];
        System.arraycopy(data, offset, padded, 0, length);

        return cipher.doFinal(padded);
    }

    /**
     * Encrypts data using AES-CBC without any padding.
     * Data must already be aligned to block size (16 bytes).
     *
     * @param data the data to encrypt (must be multiple of 16 bytes)
     * @param iv initialization vector (16 bytes)
     * @return encrypted data
     */
    public byte[] encryptNoPadding(byte[] data, byte[] iv) throws GeneralSecurityException {
        if (data.length % 16 != 0) {
            throw new IllegalArgumentException("Data must be aligned to 16 bytes");
        }
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.ENCRYPT_MODE, keySpec, ivSpec);
        return cipher.doFinal(data);
    }

    /**
     * Encrypts data from input ByteBuffer to output ByteBuffer using AES-CBC.
     * Input is padded to block size (16 bytes).
     *
     * @param input source buffer (position to limit will be read)
     * @param output destination buffer (must have enough remaining space)
     * @param iv initialization vector (16 bytes)
     * @return number of bytes written to output
     */
    public int encrypt(ByteBuffer input, ByteBuffer output, byte[] iv) throws GeneralSecurityException {
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.ENCRYPT_MODE, keySpec, ivSpec);

        int length = input.remaining();
        int blockSize = 16;
        int paddedLength = ((length + blockSize - 1) / blockSize) * blockSize;

        // If already aligned, use direct operation
        if (length == paddedLength && length > 0) {
            return cipher.doFinal(input, output);
        }

        // Need padding - create padded Direct buffer for Cipher efficiency
        ByteBuffer padded = ByteBuffer.allocateDirect(paddedLength);
        padded.put(input);
        // Set position to 0, keep limit at paddedLength (includes zero-padding)
        padded.position(0);

        return cipher.doFinal(padded, output);
    }

    /**
     * Decrypts data using AES-CBC.
     *
     * @param data the data to decrypt
     * @param iv initialization vector (16 bytes)
     * @return decrypted data
     */
    public byte[] decrypt(byte[] data, byte[] iv) throws GeneralSecurityException {
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);
        return cipher.doFinal(data);
    }

    /**
     * Decrypts a portion of data using AES-CBC.
     *
     * @param data the data to decrypt
     * @param offset start offset in data
     * @param length number of bytes to decrypt
     * @param iv initialization vector (16 bytes)
     * @return decrypted data
     */
    public byte[] decrypt(byte[] data, int offset, int length, byte[] iv) throws GeneralSecurityException {
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);
        return cipher.doFinal(data, offset, length);
    }

    /**
     * Decrypts data from input ByteBuffer to output ByteBuffer using AES-CBC.
     *
     * @param input source buffer containing encrypted data
     * @param output destination buffer
     * @param iv initialization vector (16 bytes)
     * @return number of bytes written to output
     */
    public int decrypt(ByteBuffer input, ByteBuffer output, byte[] iv) throws GeneralSecurityException {
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);
        return cipher.doFinal(input, output);
    }
}
