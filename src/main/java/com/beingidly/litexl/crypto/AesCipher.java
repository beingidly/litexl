package com.beingidly.litexl.crypto;

import javax.crypto.Cipher;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import java.nio.ByteBuffer;
import java.security.GeneralSecurityException;

/**
 * AES encryption/decryption utilities.
 */
public final class AesCipher {

    private AesCipher() {}

    /**
     * Encrypts data using AES-CBC with PKCS5 padding.
     */
    public static byte[] encrypt(byte[] data, byte[] key, byte[] iv) throws GeneralSecurityException {
        return encrypt(data, 0, data.length, key, iv);
    }

    /**
     * Encrypts a portion of data using AES-CBC with zero padding to block size.
     */
    public static byte[] encrypt(byte[] data, int offset, int length, byte[] key, byte[] iv) throws GeneralSecurityException {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
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
     * Decrypts data using AES-CBC.
     */
    public static byte[] decrypt(byte[] data, byte[] key, byte[] iv) throws GeneralSecurityException {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);
        return cipher.doFinal(data);
    }

    /**
     * Decrypts a portion of data using AES-CBC.
     */
    public static byte[] decrypt(byte[] data, int offset, int length, byte[] key, byte[] iv) throws GeneralSecurityException {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);
        return cipher.doFinal(data, offset, length);
    }

    /**
     * Encrypts data using AES-CBC with a pre-initialized Cipher instance.
     * This avoids the overhead of Cipher.getInstance() on each call.
     *
     * @param cipher a Cipher instance (will be initialized with ENCRYPT_MODE)
     * @param data the data to encrypt
     * @param key AES key
     * @param iv initialization vector
     * @return encrypted data
     */
    public static byte[] encrypt(Cipher cipher, byte[] data, byte[] key, byte[] iv) throws GeneralSecurityException {
        return encrypt(cipher, data, 0, data.length, key, iv);
    }

    /**
     * Encrypts a portion of data using AES-CBC with a pre-initialized Cipher instance.
     * This avoids the overhead of Cipher.getInstance() on each call.
     *
     * @param cipher a Cipher instance (will be initialized with ENCRYPT_MODE)
     * @param data the data to encrypt
     * @param offset start offset in data
     * @param length number of bytes to encrypt
     * @param key AES key
     * @param iv initialization vector
     * @return encrypted data
     */
    public static byte[] encrypt(Cipher cipher, byte[] data, int offset, int length, byte[] key, byte[] iv)
            throws GeneralSecurityException {
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
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
     * Decrypts data using AES-CBC with a pre-initialized Cipher instance.
     * This avoids the overhead of Cipher.getInstance() on each call.
     *
     * @param cipher a Cipher instance (will be initialized with DECRYPT_MODE)
     * @param data the data to decrypt
     * @param key AES key
     * @param iv initialization vector
     * @return decrypted data
     */
    public static byte[] decrypt(Cipher cipher, byte[] data, byte[] key, byte[] iv) throws GeneralSecurityException {
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);
        return cipher.doFinal(data);
    }

    /**
     * Decrypts a portion of data using AES-CBC with a pre-initialized Cipher instance.
     * This avoids the overhead of Cipher.getInstance() on each call.
     *
     * @param cipher a Cipher instance (will be initialized with DECRYPT_MODE)
     * @param data the data to decrypt
     * @param offset start offset in data
     * @param length number of bytes to decrypt
     * @param key AES key
     * @param iv initialization vector
     * @return decrypted data
     */
    public static byte[] decrypt(Cipher cipher, byte[] data, int offset, int length, byte[] key, byte[] iv)
            throws GeneralSecurityException {
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);
        return cipher.doFinal(data, offset, length);
    }

    /**
     * Encrypts data from input ByteBuffer to output ByteBuffer using AES-CBC with a pre-initialized Cipher instance.
     * Input is padded to block size (16 bytes).
     * This avoids the overhead of Cipher.getInstance() on each call.
     *
     * @param cipher a Cipher instance (will be initialized with ENCRYPT_MODE)
     * @param input source buffer (position to limit will be read)
     * @param output destination buffer (must have enough remaining space)
     * @param key AES key
     * @param iv initialization vector
     * @return number of bytes written to output
     */
    public static int encrypt(Cipher cipher, ByteBuffer input, ByteBuffer output, byte[] key, byte[] iv)
            throws GeneralSecurityException {
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
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
     * Decrypts data from input ByteBuffer to output ByteBuffer using AES-CBC with a pre-initialized Cipher instance.
     * This avoids the overhead of Cipher.getInstance() on each call.
     *
     * @param cipher a Cipher instance (will be initialized with DECRYPT_MODE)
     * @param input source buffer containing encrypted data
     * @param output destination buffer
     * @param key AES key
     * @param iv initialization vector
     * @return number of bytes written to output
     */
    public static int decrypt(Cipher cipher, ByteBuffer input, ByteBuffer output, byte[] key, byte[] iv)
            throws GeneralSecurityException {
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);

        return cipher.doFinal(input, output);
    }

    /**
     * Encrypts data using AES-CBC without any padding.
     * Data must already be aligned to block size (16 bytes).
     */
    public static byte[] encryptNoPadding(byte[] data, byte[] key, byte[] iv) throws GeneralSecurityException {
        if (data.length % 16 != 0) {
            throw new IllegalArgumentException("Data must be aligned to 16 bytes");
        }
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.ENCRYPT_MODE, keySpec, ivSpec);
        return cipher.doFinal(data);
    }

    /**
     * Encrypts data from input ByteBuffer to output ByteBuffer using AES-CBC.
     * Input is padded to block size (16 bytes).
     * Uses Cipher.doFinal(ByteBuffer, ByteBuffer) for efficient operation.
     *
     * @param input source buffer (position to limit will be read)
     * @param output destination buffer (must have enough remaining space)
     * @param key AES key (16, 24, or 32 bytes)
     * @param iv initialization vector (16 bytes)
     * @return number of bytes written to output
     */
    public static int encrypt(ByteBuffer input, ByteBuffer output, byte[] key, byte[] iv)
            throws GeneralSecurityException {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
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
     * Decrypts data from input ByteBuffer to output ByteBuffer using AES-CBC.
     * Uses Cipher.doFinal(ByteBuffer, ByteBuffer) for zero-copy operation.
     *
     * @param input source buffer containing encrypted data
     * @param output destination buffer
     * @param key AES key
     * @param iv initialization vector
     * @return number of bytes written to output
     */
    public static int decrypt(ByteBuffer input, ByteBuffer output, byte[] key, byte[] iv)
            throws GeneralSecurityException {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");
        SecretKeySpec keySpec = new SecretKeySpec(key, "AES");
        IvParameterSpec ivSpec = new IvParameterSpec(iv);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);

        return cipher.doFinal(input, output);
    }
}
