package com.beingidly.litexl.crypto;

import java.io.*;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.util.Arrays;
import javax.crypto.Cipher;

/**
 * ECMA-376 Agile Encryption encryptor.
 *
 * <p>Usage:</p>
 * <pre>{@code
 * AgileEncryptor encryptor = new AgileEncryptor(options);
 * EncryptionResult result = encryptor.encrypt(ByteBuffer.wrap(data));
 * }</pre>
 */
public final class AgileEncryptor {

    private static final int SALT_SIZE = 16;
    private static final int BLOCK_SIZE = 16;
    private static final int SEGMENT_SIZE = 4096;

    private final EncryptionOptions options;
    private final SecureRandom random = new SecureRandom();

    // Reusable crypto instances (allocated once per encryptor)
    private final Cipher cipher;
    private final MessageDigest sha512;
    private final ByteBuffer ivIndexBuf;

    /**
     * Creates a new encryptor with the given options.
     *
     * @param options the encryption options
     */
    public AgileEncryptor(EncryptionOptions options) {
        this.options = options;
        try {
            this.cipher = Cipher.getInstance("AES/CBC/NoPadding");
            this.sha512 = MessageDigest.getInstance("SHA-512");
        } catch (GeneralSecurityException e) {
            throw new RuntimeException("Required crypto algorithms not available", e);
        }
        this.ivIndexBuf = ByteBuffer.allocate(4).order(ByteOrder.LITTLE_ENDIAN);
    }

    /**
     * Encrypts data using ECMA-376 Agile Encryption.
     *
     * @param data the data to encrypt as ByteBuffer
     * @return the encryption result containing encrypted data and metadata
     * @throws GeneralSecurityException if encryption fails
     */
    public EncryptionResult encrypt(ByteBuffer data) throws GeneralSecurityException {
        int keyBits = options.algorithm() == EncryptionOptions.Algorithm.AES_256 ? 256 : 128;

        // Generate salt and encryption key
        byte[] salt = new byte[SALT_SIZE];
        random.nextBytes(salt);

        byte[] encryptionKey = new byte[keyBits / 8];
        random.nextBytes(encryptionKey);

        // Derive intermediate hash once (expensive spinCount iterations)
        byte[] intermediateHash = KeyDerivation.deriveIntermediateHash(
            options.password(), salt, options.spinCount()
        );

        // Derive all three keys from intermediate (fast, no iterations)
        byte[] keyDerivedKey = KeyDerivation.deriveKeyFromIntermediate(
            intermediateHash, keyBits, KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );
        byte[] verifierInputKey = KeyDerivation.deriveKeyFromIntermediate(
            intermediateHash, keyBits, KeyDerivation.BLOCK_KEY_VERIFIER_INPUT
        );
        byte[] verifierHashKey = KeyDerivation.deriveKeyFromIntermediate(
            intermediateHash, keyBits, KeyDerivation.BLOCK_KEY_VERIFIER_VALUE
        );

        // Encrypt the encryption key (small data - use byte[] API with reusable cipher)
        byte[] iv = new byte[BLOCK_SIZE];
        random.nextBytes(iv);
        byte[] encryptedKey = AesCipher.encrypt(cipher, encryptionKey, keyDerivedKey, iv);

        // Generate verifier (small data - use byte[] API with reusable cipher)
        byte[] verifierInput = new byte[SALT_SIZE];
        random.nextBytes(verifierInput);
        byte[] encryptedVerifierInput = AesCipher.encrypt(cipher, verifierInput, verifierInputKey, iv);

        sha512.reset();
        byte[] verifierHash = sha512.digest(verifierInput);

        byte[] encryptedVerifierHash = AesCipher.encrypt(
            cipher,
            Arrays.copyOf(verifierHash, 32),
            verifierHashKey, iv
        );

        // Encrypt the data (large data - use ByteBuffer API)
        ByteBuffer encryptedData = encryptData(data, encryptionKey);

        // Build encryption info XML
        String encryptionInfo = buildEncryptionInfoXml(
            keyBits, salt, options.spinCount(),
            encryptedKey, encryptedVerifierInput, encryptedVerifierHash
        );

        // Combine into output
        byte[] infoBytes = encryptionInfo.getBytes(StandardCharsets.UTF_8);
        int totalSize = 8 + infoBytes.length + encryptedData.remaining();
        ByteBuffer output = ByteBuffer.allocateDirect(totalSize).order(ByteOrder.BIG_ENDIAN);

        // Version header
        output.putShort((short) 0x0004); // Major version 4
        output.putShort((short) 0x0004); // Minor version 4
        output.putInt(0x00040); // Flags: Agile

        output.put(infoBytes);
        output.put(encryptedData);
        output.flip();

        return new EncryptionResult(
            output,
            encryptedData,
            encryptionKey,
            salt
        );
    }

    private ByteBuffer encryptData(ByteBuffer data, byte[] key) throws GeneralSecurityException {
        int dataLength = data.remaining();

        // Calculate output size: 8-byte header + padded segments
        int numSegments = (dataLength + SEGMENT_SIZE - 1) / SEGMENT_SIZE;
        int maxOutputSize = 8 + numSegments * SEGMENT_SIZE;

        ByteBuffer output = ByteBuffer.allocateDirect(maxOutputSize).order(ByteOrder.LITTLE_ENDIAN);

        // Write original size
        output.putLong(dataLength);

        // Reusable Direct buffers for encryption
        ByteBuffer inputBuf = ByteBuffer.allocateDirect(SEGMENT_SIZE);
        ByteBuffer encryptBuf = ByteBuffer.allocateDirect(SEGMENT_SIZE);

        int segmentIndex = 0;
        while (data.hasRemaining()) {
            int segmentLen = Math.min(SEGMENT_SIZE, data.remaining());

            // Copy segment to input buffer
            inputBuf.clear();
            int oldLimit = data.limit();
            data.limit(data.position() + segmentLen);
            inputBuf.put(data);
            data.limit(oldLimit);
            inputBuf.flip();

            // Generate IV and encrypt
            byte[] segmentIv = generateSegmentIv(segmentIndex, key);
            encryptBuf.clear();
            AesCipher.encrypt(cipher, inputBuf, encryptBuf, key, segmentIv);

            // Copy to output
            encryptBuf.flip();
            output.put(encryptBuf);

            segmentIndex++;
        }

        output.flip();
        return output;
    }

    private byte[] generateSegmentIv(int segmentIndex, byte[] salt) {
        // IV = SHA-512(salt + LE32(segmentIndex))[0:16]
        sha512.reset();
        sha512.update(salt);
        ivIndexBuf.clear();
        ivIndexBuf.putInt(segmentIndex);
        ivIndexBuf.flip();
        sha512.update(ivIndexBuf);
        byte[] hash = sha512.digest();
        return Arrays.copyOf(hash, 16);
    }

    private String buildEncryptionInfoXml(
            int keyBits, byte[] salt, int spinCount,
            byte[] encryptedKey,
            byte[] encryptedVerifierInput, byte[] encryptedVerifierHash) {

        String saltB64 = java.util.Base64.getEncoder().encodeToString(salt);
        String encKeyB64 = java.util.Base64.getEncoder().encodeToString(encryptedKey);
        String encVerInputB64 = java.util.Base64.getEncoder().encodeToString(encryptedVerifierInput);
        String encVerHashB64 = java.util.Base64.getEncoder().encodeToString(encryptedVerifierHash);

        return String.format("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption">
              <keyData saltSize="%d" blockSize="16" keyBits="%d" hashSize="64"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       hashAlgorithm="SHA512" saltValue="%s"/>
              <dataIntegrity encryptedHmacKey="" encryptedHmacValue=""/>
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
                                  spinCount="%d" saltSize="%d" blockSize="16" keyBits="%d"
                                  hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  hashAlgorithm="SHA512" saltValue="%s"
                                  encryptedVerifierHashInput="%s"
                                  encryptedVerifierHashValue="%s"
                                  encryptedKeyValue="%s"/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
            """,
            SALT_SIZE, keyBits, saltB64,
            spinCount, SALT_SIZE, keyBits, saltB64,
            encVerInputB64, encVerHashB64, encKeyB64
        );
    }

    /**
     * Result of encryption containing all necessary components.
     */
    public record EncryptionResult(
        ByteBuffer encryptionInfoWithData,
        ByteBuffer encryptedData,
        byte[] encryptionKey,
        byte[] keyDataSalt
    ) {}
}
