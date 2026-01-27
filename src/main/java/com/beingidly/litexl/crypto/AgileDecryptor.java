package com.beingidly.litexl.crypto;

import com.beingidly.litexl.CorruptFileException;
import com.beingidly.litexl.InvalidPasswordException;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import java.util.Arrays;
import java.util.Base64;
import javax.crypto.Cipher;
import javax.xml.stream.*;
import org.jspecify.annotations.Nullable;

/**
 * ECMA-376 Agile Encryption decryptor.
 *
 * <p>Usage:</p>
 * <pre>{@code
 * AgileDecryptor decryptor = new AgileDecryptor(password);
 * decryptor.parseEncryptionInfo(encryptionInfoData, encryptedPackage);
 * ByteBuffer decrypted = decryptor.decrypt();
 * }</pre>
 */
public final class AgileDecryptor {

    private final String password;
    private @Nullable EncryptionInfo parsedInfo;

    // Reusable crypto instance
    private final Cipher cipher;

    /**
     * Creates a new decryptor with the given password.
     *
     * @param password the password for decryption
     */
    public AgileDecryptor(String password) {
        this.password = password;
        try {
            this.cipher = Cipher.getInstance("AES/CBC/NoPadding");
        } catch (GeneralSecurityException e) {
            throw new RuntimeException("AES cipher not available", e);
        }
    }

    /**
     * Parses the EncryptionInfo stream from an encrypted document.
     *
     * @param encryptionInfoData the raw EncryptionInfo stream data
     * @param encryptedPackage the encrypted package data as ByteBuffer
     * @throws IOException if parsing fails
     */
    public void parseEncryptionInfo(ByteBuffer encryptionInfoData, ByteBuffer encryptedPackage) throws IOException {
        ByteBuffer orderedData = encryptionInfoData.order(ByteOrder.LITTLE_ENDIAN);

        // Read version header
        short majorVersion = orderedData.getShort();
        short minorVersion = orderedData.getShort();
        orderedData.getInt(); // flags - reserved for future use

        if (majorVersion != 4 || minorVersion != 4) {
            throw new CorruptFileException("Unsupported encryption version: " + majorVersion + "." + minorVersion);
        }

        // Rest is XML
        byte[] xmlData = new byte[orderedData.remaining()];
        orderedData.get(xmlData);

        this.parsedInfo = parseAgileXml(xmlData, encryptedPackage);
    }

    private EncryptionInfo parseAgileXml(byte[] xmlData, ByteBuffer encryptedPackage) throws IOException {
        try {
            XMLInputFactory factory = XMLInputFactory.newInstance();
            factory.setProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES, false);
            factory.setProperty(XMLInputFactory.SUPPORT_DTD, false);

            XMLStreamReader reader = factory.createXMLStreamReader(new ByteArrayInputStream(xmlData));

            int keyBits = 256;
            byte[] keyDataSalt = null;
            byte[] encryptedKeySalt = null;
            int spinCount = 100000;
            byte[] encryptedKey = null;
            byte[] encryptedVerifierInput = null;
            byte[] encryptedVerifierHash = null;

            while (reader.hasNext()) {
                int event = reader.next();
                if (event == XMLStreamConstants.START_ELEMENT) {
                    String localName = reader.getLocalName();

                    if ("keyData".equals(localName)) {
                        String keyBitsStr = reader.getAttributeValue(null, "keyBits");
                        if (keyBitsStr != null) {
                            keyBits = Integer.parseInt(keyBitsStr);
                        }
                        String saltValue = reader.getAttributeValue(null, "saltValue");
                        if (saltValue != null) {
                            keyDataSalt = Base64.getDecoder().decode(saltValue);
                        }
                    } else if ("encryptedKey".equals(localName)) {
                        String spinCountStr = reader.getAttributeValue(null, "spinCount");
                        if (spinCountStr != null) {
                            spinCount = Integer.parseInt(spinCountStr);
                        }

                        String saltValue = reader.getAttributeValue(null, "saltValue");
                        if (saltValue != null) {
                            encryptedKeySalt = Base64.getDecoder().decode(saltValue);
                        }

                        String encVerInput = reader.getAttributeValue(null, "encryptedVerifierHashInput");
                        if (encVerInput != null) {
                            encryptedVerifierInput = Base64.getDecoder().decode(encVerInput);
                        }

                        String encVerHash = reader.getAttributeValue(null, "encryptedVerifierHashValue");
                        if (encVerHash != null) {
                            encryptedVerifierHash = Base64.getDecoder().decode(encVerHash);
                        }

                        String encKeyValue = reader.getAttributeValue(null, "encryptedKeyValue");
                        if (encKeyValue != null) {
                            encryptedKey = Base64.getDecoder().decode(encKeyValue);
                        }
                    }
                }
            }

            reader.close();

            // Use encryptedKeySalt for password verification, fall back to keyDataSalt
            byte[] salt = encryptedKeySalt != null ? encryptedKeySalt : keyDataSalt;
            // Use keyDataSalt for data decryption IV, fall back to salt
            byte[] dataSalt = keyDataSalt != null ? keyDataSalt : salt;

            if (salt == null || encryptedKey == null || encryptedVerifierInput == null || encryptedVerifierHash == null) {
                throw new CorruptFileException("Missing encryption parameters");
            }

            // IV for verifier decryption is the encryptedKeySalt
            byte[] iv = new byte[16];
            System.arraycopy(salt, 0, iv, 0, Math.min(salt.length, 16));

            return new EncryptionInfo(
                keyBits,
                salt,
                dataSalt,
                spinCount,
                iv,
                encryptedKey,
                encryptedVerifierInput,
                encryptedVerifierHash,
                encryptedPackage
            );

        } catch (XMLStreamException e) {
            throw new IOException("Failed to parse encryption info XML", e);
        }
    }

    /**
     * Decrypts the encrypted package.
     *
     * @return the decrypted data as a ByteBuffer
     * @throws GeneralSecurityException if decryption fails
     * @throws IllegalStateException if parseEncryptionInfo was not called first
     */
    public ByteBuffer decrypt() throws GeneralSecurityException {
        if (parsedInfo == null) {
            throw new IllegalStateException("parseEncryptionInfo must be called first");
        }

        int keyBits = parsedInfo.keyBits();

        // Derive intermediate hash once (expensive spinCount iterations)
        byte[] intermediateHash = KeyDerivation.deriveIntermediateHash(
            password, parsedInfo.salt(), parsedInfo.spinCount()
        );

        // Derive keys from intermediate hash (fast)
        byte[] verifierInputKey = KeyDerivation.deriveKeyFromIntermediate(
            intermediateHash, keyBits, KeyDerivation.BLOCK_KEY_VERIFIER_INPUT
        );

        // Decrypt verifier input
        byte[] decryptedVerifierInput = AesCipher.decrypt(
            cipher, parsedInfo.encryptedVerifierInput(), verifierInputKey, parsedInfo.iv()
        );

        // Verify password by checking verifier hash
        byte[] verifierHashKey = KeyDerivation.deriveKeyFromIntermediate(
            intermediateHash, keyBits, KeyDerivation.BLOCK_KEY_VERIFIER_VALUE
        );

        byte[] decryptedVerifierHash = AesCipher.decrypt(
            cipher, parsedInfo.encryptedVerifierHash(), verifierHashKey, parsedInfo.iv()
        );

        // Compute expected hash
        try {
            MessageDigest sha512 = MessageDigest.getInstance("SHA-512");
            byte[] expectedHash = sha512.digest(decryptedVerifierInput);

            if (!Arrays.equals(Arrays.copyOf(expectedHash, 32), Arrays.copyOf(decryptedVerifierHash, 32))) {
                throw new InvalidPasswordException();
            }
        } catch (InvalidPasswordException e) {
            throw e;
        } catch (Exception e) {
            throw new InvalidPasswordException();
        }

        // Decrypt the encryption key
        byte[] keyDerivedKey = KeyDerivation.deriveKeyFromIntermediate(
            intermediateHash, keyBits, KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        byte[] encryptionKey = AesCipher.decrypt(
            cipher, parsedInfo.encryptedKey(), keyDerivedKey, parsedInfo.iv()
        );
        encryptionKey = Arrays.copyOf(encryptionKey, keyBits / 8);

        // Decrypt the data using keyDataSalt for IV generation
        return decryptData(parsedInfo.encryptedData(), encryptionKey, parsedInfo.keyDataSalt());
    }

    private ByteBuffer decryptData(ByteBuffer data, byte[] key, byte[] salt) throws GeneralSecurityException {
        ByteBuffer inputBuf = data.order(ByteOrder.LITTLE_ENDIAN);

        // Read original size
        long originalSize = inputBuf.getLong();
        if (originalSize < 0 || originalSize > Integer.MAX_VALUE) {
            throw new GeneralSecurityException("Invalid data size");
        }

        // Pre-allocate output buffer with Direct ByteBuffer for efficiency
        ByteBuffer outputBuf = ByteBuffer.allocateDirect((int) originalSize);

        int segmentSize = 4096;
        int segmentIndex = 0;
        int remaining = (int) originalSize;

        // Reusable Direct buffers for decryption
        int maxPaddedLen = ((segmentSize + 15) / 16) * 16;
        ByteBuffer segmentInputBuf = ByteBuffer.allocateDirect(maxPaddedLen);
        ByteBuffer decryptBuf = ByteBuffer.allocateDirect(maxPaddedLen);
        MessageDigest sha512 = MessageDigest.getInstance("SHA-512");
        ByteBuffer indexBuf = ByteBuffer.allocate(4).order(ByteOrder.LITTLE_ENDIAN);
        byte[] ivBuffer = new byte[16];

        while (inputBuf.hasRemaining() && remaining > 0) {
            int encryptedSegmentLen = Math.min(segmentSize, remaining);
            int paddedLen = ((encryptedSegmentLen + 15) / 16) * 16;

            int toRead = Math.min(paddedLen, inputBuf.remaining());

            // Copy segment to direct buffer
            segmentInputBuf.clear();
            int oldLimit = inputBuf.limit();
            inputBuf.limit(inputBuf.position() + toRead);
            segmentInputBuf.put(inputBuf);
            inputBuf.limit(oldLimit);
            segmentInputBuf.flip();
            segmentInputBuf.limit(paddedLen);

            // Generate IV: SHA-512(salt + LE32(segmentIndex))[0:16] - reuse digest
            sha512.reset();
            sha512.update(salt);
            indexBuf.clear();
            indexBuf.putInt(segmentIndex);
            indexBuf.flip();
            sha512.update(indexBuf);
            byte[] hash = sha512.digest();
            System.arraycopy(hash, 0, ivBuffer, 0, 16);

            decryptBuf.clear();
            AesCipher.decrypt(cipher, segmentInputBuf, decryptBuf, key, ivBuffer);
            decryptBuf.flip();

            decryptBuf.limit(encryptedSegmentLen);
            outputBuf.put(decryptBuf);

            remaining -= encryptedSegmentLen;
            segmentIndex++;
        }

        outputBuf.flip();
        return outputBuf;
    }

    /**
     * Parsed encryption info from an encrypted file.
     */
    private record EncryptionInfo(
        int keyBits,
        byte[] salt,
        byte[] keyDataSalt,
        int spinCount,
        byte[] iv,
        byte[] encryptedKey,
        byte[] encryptedVerifierInput,
        byte[] encryptedVerifierHash,
        ByteBuffer encryptedData
    ) {}
}
