package com.beingidly.litexl.crypto;

import org.junit.jupiter.api.Test;

import javax.crypto.Cipher;
import java.nio.ByteBuffer;
import java.nio.charset.StandardCharsets;
import java.security.GeneralSecurityException;
import java.util.Arrays;

import static org.junit.jupiter.api.Assertions.*;

class AesCipherTest {

    @Test
    void encryptDecrypt128() throws GeneralSecurityException {
        byte[] key = new byte[16]; // 128-bit key
        Arrays.fill(key, (byte) 0x42);

        byte[] iv = new byte[16];
        Arrays.fill(iv, (byte) 0x24);

        byte[] plaintext = "Hello, World!!!".getBytes(); // 16 bytes (1 block)

        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);
        byte[] decrypted = AesCipher.decrypt(ciphertext, key, iv);

        assertArrayEquals(plaintext, Arrays.copyOf(decrypted, plaintext.length));
    }

    @Test
    void encryptDecrypt256() throws GeneralSecurityException {
        byte[] key = new byte[32]; // 256-bit key
        Arrays.fill(key, (byte) 0x42);

        byte[] iv = new byte[16];
        Arrays.fill(iv, (byte) 0x24);

        byte[] plaintext = "Hello, World!!!".getBytes();

        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);
        byte[] decrypted = AesCipher.decrypt(ciphertext, key, iv);

        assertArrayEquals(plaintext, Arrays.copyOf(decrypted, plaintext.length));
    }

    @Test
    void encryptPadsToBlockSize() throws GeneralSecurityException {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];

        byte[] plaintext = "Short".getBytes(); // Less than 16 bytes

        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);

        // Ciphertext should be padded to 16 bytes (one block)
        assertEquals(16, ciphertext.length);
    }

    @Test
    void encryptMultipleBlocks() throws GeneralSecurityException {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];

        byte[] plaintext = new byte[33]; // > 2 blocks
        Arrays.fill(plaintext, (byte) 'A');

        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);

        // Should be 48 bytes (3 blocks)
        assertEquals(48, ciphertext.length);
    }

    @Test
    void differentKeysProduceDifferentCiphertext() throws GeneralSecurityException {
        byte[] key1 = new byte[16];
        byte[] key2 = new byte[16];
        key2[0] = 1;

        byte[] iv = new byte[16];
        byte[] plaintext = "Test message".getBytes();

        byte[] ciphertext1 = AesCipher.encrypt(plaintext, key1, iv);
        byte[] ciphertext2 = AesCipher.encrypt(plaintext, key2, iv);

        assertFalse(Arrays.equals(ciphertext1, ciphertext2));
    }

    @Test
    void differentIvsProduceDifferentCiphertext() throws GeneralSecurityException {
        byte[] key = new byte[16];

        byte[] iv1 = new byte[16];
        byte[] iv2 = new byte[16];
        iv2[0] = 1;

        byte[] plaintext = "Test message".getBytes();

        byte[] ciphertext1 = AesCipher.encrypt(plaintext, key, iv1);
        byte[] ciphertext2 = AesCipher.encrypt(plaintext, key, iv2);

        assertFalse(Arrays.equals(ciphertext1, ciphertext2));
    }

    @Test
    void invalidKeySizeThrows() {
        byte[] key = new byte[15]; // Invalid size
        byte[] iv = new byte[16];
        byte[] data = "Test".getBytes();

        assertThrows(GeneralSecurityException.class, () ->
            AesCipher.encrypt(data, key, iv)
        );
    }

    @Test
    void invalidIvSizeThrows() {
        byte[] key = new byte[16];
        byte[] iv = new byte[15]; // Invalid size
        byte[] data = "Test".getBytes();

        assertThrows(GeneralSecurityException.class, () ->
            AesCipher.encrypt(data, key, iv)
        );
    }

    @Test
    void encryptWithByteBuffer() throws Exception {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x01);
        Arrays.fill(iv, (byte) 0x02);

        byte[] plaintext = "Hello, World!!!X".getBytes(StandardCharsets.UTF_8); // 16 bytes
        ByteBuffer input = ByteBuffer.wrap(plaintext);
        ByteBuffer output = ByteBuffer.allocate(32);

        int written = AesCipher.encrypt(input, output, key, iv);

        assertTrue(written > 0);
        assertEquals(16, written);
        output.flip();

        // Verify by decrypting
        byte[] encrypted = new byte[written];
        output.get(encrypted);
        byte[] decrypted = AesCipher.decrypt(encrypted, key, iv);
        assertArrayEquals(plaintext, Arrays.copyOf(decrypted, plaintext.length));
    }

    @Test
    void decryptWithByteBuffer() throws Exception {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x01);
        Arrays.fill(iv, (byte) 0x02);

        byte[] plaintext = "Hello, World!!!X".getBytes(StandardCharsets.UTF_8);
        byte[] encrypted = AesCipher.encrypt(plaintext, key, iv);

        ByteBuffer input = ByteBuffer.wrap(encrypted);
        ByteBuffer output = ByteBuffer.allocate(encrypted.length);

        int written = AesCipher.decrypt(input, output, key, iv);

        assertEquals(encrypted.length, written);
        output.flip();
        byte[] decrypted = new byte[written];
        output.get(decrypted);
        assertArrayEquals(plaintext, Arrays.copyOf(decrypted, plaintext.length));
    }

    @Test
    void encryptDecryptWithReusedCipher() throws Exception {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");

        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x42);
        Arrays.fill(iv, (byte) 0x24);

        // First encrypt/decrypt cycle
        byte[] plaintext1 = "Hello, World!!!".getBytes(StandardCharsets.UTF_8);
        byte[] ciphertext1 = AesCipher.encrypt(cipher, plaintext1, key, iv);
        byte[] decrypted1 = AesCipher.decrypt(cipher, ciphertext1, key, iv);
        assertArrayEquals(plaintext1, Arrays.copyOf(decrypted1, plaintext1.length));

        // Second encrypt/decrypt cycle with same Cipher instance
        byte[] plaintext2 = "Goodbye, World!".getBytes(StandardCharsets.UTF_8);
        byte[] ciphertext2 = AesCipher.encrypt(cipher, plaintext2, key, iv);
        byte[] decrypted2 = AesCipher.decrypt(cipher, ciphertext2, key, iv);
        assertArrayEquals(plaintext2, Arrays.copyOf(decrypted2, plaintext2.length));

        // Verify results match non-reused Cipher methods
        byte[] expected1 = AesCipher.encrypt(plaintext1, key, iv);
        byte[] expected2 = AesCipher.encrypt(plaintext2, key, iv);
        assertArrayEquals(expected1, ciphertext1);
        assertArrayEquals(expected2, ciphertext2);
    }

    @Test
    void encryptDecryptWithReusedCipherByteBuffer() throws Exception {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");

        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x01);
        Arrays.fill(iv, (byte) 0x02);

        // First encrypt/decrypt cycle with ByteBuffer
        byte[] plaintext1 = "Hello, World!!!X".getBytes(StandardCharsets.UTF_8); // 16 bytes
        ByteBuffer input1 = ByteBuffer.wrap(plaintext1);
        ByteBuffer output1 = ByteBuffer.allocate(32);
        int written1 = AesCipher.encrypt(cipher, input1, output1, key, iv);
        assertEquals(16, written1);
        output1.flip();

        ByteBuffer decryptInput1 = ByteBuffer.allocate(written1);
        decryptInput1.put(output1);
        decryptInput1.flip();
        ByteBuffer decryptOutput1 = ByteBuffer.allocate(32);
        int decrypted1 = AesCipher.decrypt(cipher, decryptInput1, decryptOutput1, key, iv);
        decryptOutput1.flip();
        byte[] result1 = new byte[decrypted1];
        decryptOutput1.get(result1);
        assertArrayEquals(plaintext1, Arrays.copyOf(result1, plaintext1.length));

        // Second encrypt/decrypt cycle with same Cipher instance
        byte[] plaintext2 = "Goodbye, World!!".getBytes(StandardCharsets.UTF_8); // 16 bytes
        ByteBuffer input2 = ByteBuffer.wrap(plaintext2);
        ByteBuffer output2 = ByteBuffer.allocate(32);
        int written2 = AesCipher.encrypt(cipher, input2, output2, key, iv);
        assertEquals(16, written2);
        output2.flip();

        ByteBuffer decryptInput2 = ByteBuffer.allocate(written2);
        decryptInput2.put(output2);
        decryptInput2.flip();
        ByteBuffer decryptOutput2 = ByteBuffer.allocate(32);
        int decrypted2 = AesCipher.decrypt(cipher, decryptInput2, decryptOutput2, key, iv);
        decryptOutput2.flip();
        byte[] result2 = new byte[decrypted2];
        decryptOutput2.get(result2);
        assertArrayEquals(plaintext2, Arrays.copyOf(result2, plaintext2.length));
    }

    @Test
    void encryptDecryptWithReusedCipherOffsetLength() throws Exception {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");

        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x42);
        Arrays.fill(iv, (byte) 0x24);

        // Data with offset
        byte[] data = "PREFIXHello, World!!!SUFFIX".getBytes(StandardCharsets.UTF_8);
        int offset = 6; // Skip "PREFIX"
        int length = 15; // "Hello, World!!!"

        // Encrypt with offset/length using reused Cipher
        byte[] ciphertext = AesCipher.encrypt(cipher, data, offset, length, key, iv);

        // Decrypt using reused Cipher
        byte[] decrypted = AesCipher.decrypt(cipher, ciphertext, key, iv);

        // Verify
        byte[] expected = new byte[length];
        System.arraycopy(data, offset, expected, 0, length);
        assertArrayEquals(expected, Arrays.copyOf(decrypted, length));

        // Verify matches non-reused method
        byte[] expectedCiphertext = AesCipher.encrypt(data, offset, length, key, iv);
        assertArrayEquals(expectedCiphertext, ciphertext);
    }

    @Test
    void encryptDecrypt192() throws GeneralSecurityException {
        byte[] key = new byte[24]; // 192-bit key
        Arrays.fill(key, (byte) 0x42);

        byte[] iv = new byte[16];
        Arrays.fill(iv, (byte) 0x24);

        byte[] plaintext = "Hello, World!!!".getBytes(StandardCharsets.UTF_8);

        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);
        byte[] decrypted = AesCipher.decrypt(ciphertext, key, iv);

        assertArrayEquals(plaintext, Arrays.copyOf(decrypted, plaintext.length));
    }

    @Test
    void decryptWithOffsetLength() throws GeneralSecurityException {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x42);
        Arrays.fill(iv, (byte) 0x24);

        byte[] plaintext = "Hello, World!!!X".getBytes(StandardCharsets.UTF_8); // 16 bytes
        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);

        // Embed ciphertext in larger array
        byte[] paddedCiphertext = new byte[ciphertext.length + 20];
        System.arraycopy(ciphertext, 0, paddedCiphertext, 10, ciphertext.length);

        // Decrypt with offset/length
        byte[] decrypted = AesCipher.decrypt(paddedCiphertext, 10, ciphertext.length, key, iv);

        assertArrayEquals(plaintext, Arrays.copyOf(decrypted, plaintext.length));
    }

    @Test
    void decryptWithReusedCipherOffsetLength() throws Exception {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");

        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x42);
        Arrays.fill(iv, (byte) 0x24);

        byte[] plaintext = "Hello, World!!!X".getBytes(StandardCharsets.UTF_8); // 16 bytes
        byte[] ciphertext = AesCipher.encrypt(cipher, plaintext, key, iv);

        // Embed ciphertext in larger array
        byte[] paddedCiphertext = new byte[ciphertext.length + 20];
        System.arraycopy(ciphertext, 0, paddedCiphertext, 10, ciphertext.length);

        // Decrypt with offset/length using reused Cipher
        byte[] decrypted = AesCipher.decrypt(cipher, paddedCiphertext, 10, ciphertext.length, key, iv);

        assertArrayEquals(plaintext, Arrays.copyOf(decrypted, plaintext.length));
    }

    @Test
    void encryptNoPaddingBlockAligned() throws GeneralSecurityException {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x42);
        Arrays.fill(iv, (byte) 0x24);

        // Data exactly aligned to block size (32 bytes = 2 blocks)
        byte[] plaintext = "Hello, World!!!XHello, World!!!X".getBytes(StandardCharsets.UTF_8);
        assertEquals(32, plaintext.length);

        byte[] ciphertext = AesCipher.encryptNoPadding(plaintext, key, iv);

        assertEquals(32, ciphertext.length);

        // Verify decryption
        byte[] decrypted = AesCipher.decrypt(ciphertext, key, iv);
        assertArrayEquals(plaintext, decrypted);
    }

    @Test
    void encryptNoPaddingThrowsOnUnalignedData() {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];

        byte[] plaintext = "NotAligned".getBytes(StandardCharsets.UTF_8); // 10 bytes

        assertThrows(IllegalArgumentException.class, () ->
            AesCipher.encryptNoPadding(plaintext, key, iv)
        );
    }

    @Test
    void encryptEmptyData() throws GeneralSecurityException {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];

        byte[] plaintext = new byte[0];

        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);

        // Empty data pads to 0 bytes (no block needed)
        assertEquals(0, ciphertext.length);
    }

    @Test
    void encryptExactlyOneBlock() throws GeneralSecurityException {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x42);
        Arrays.fill(iv, (byte) 0x24);

        // Exactly 16 bytes (1 block)
        byte[] plaintext = "0123456789ABCDEF".getBytes(StandardCharsets.UTF_8);
        assertEquals(16, plaintext.length);

        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);

        // Should stay at 16 bytes (no extra padding block)
        assertEquals(16, ciphertext.length);

        // Verify roundtrip
        byte[] decrypted = AesCipher.decrypt(ciphertext, key, iv);
        assertArrayEquals(plaintext, decrypted);
    }

    @Test
    void encryptExactlyTwoBlocks() throws GeneralSecurityException {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x42);
        Arrays.fill(iv, (byte) 0x24);

        // Exactly 32 bytes (2 blocks)
        byte[] plaintext = "0123456789ABCDEF0123456789ABCDEF".getBytes(StandardCharsets.UTF_8);
        assertEquals(32, plaintext.length);

        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);

        // Should stay at 32 bytes
        assertEquals(32, ciphertext.length);

        // Verify roundtrip
        byte[] decrypted = AesCipher.decrypt(ciphertext, key, iv);
        assertArrayEquals(plaintext, decrypted);
    }

    @Test
    void encryptPaddingBoundaryPlus1() throws GeneralSecurityException {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];

        // 17 bytes (1 block + 1 byte) -> should pad to 32 bytes
        byte[] plaintext = "0123456789ABCDEFX".getBytes(StandardCharsets.UTF_8);
        assertEquals(17, plaintext.length);

        byte[] ciphertext = AesCipher.encrypt(plaintext, key, iv);

        assertEquals(32, ciphertext.length);
    }

    @Test
    void encryptByteBufferWithPadding() throws Exception {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x01);
        Arrays.fill(iv, (byte) 0x02);

        // 10 bytes - needs padding to 16
        byte[] plaintext = "HelloWorld".getBytes(StandardCharsets.UTF_8);
        assertEquals(10, plaintext.length);

        ByteBuffer input = ByteBuffer.wrap(plaintext);
        ByteBuffer output = ByteBuffer.allocate(32);

        int written = AesCipher.encrypt(input, output, key, iv);

        assertEquals(16, written); // Padded to 16 bytes
        output.flip();

        // Verify by decrypting
        byte[] encrypted = new byte[written];
        output.get(encrypted);
        byte[] decrypted = AesCipher.decrypt(encrypted, key, iv);
        assertArrayEquals(plaintext, Arrays.copyOf(decrypted, plaintext.length));
    }

    @Test
    void encryptByteBufferWithReusedCipherAndPadding() throws Exception {
        Cipher cipher = Cipher.getInstance("AES/CBC/NoPadding");

        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x01);
        Arrays.fill(iv, (byte) 0x02);

        // 10 bytes - needs padding to 16
        byte[] plaintext = "HelloWorld".getBytes(StandardCharsets.UTF_8);
        assertEquals(10, plaintext.length);

        ByteBuffer input = ByteBuffer.wrap(plaintext);
        ByteBuffer output = ByteBuffer.allocate(32);

        int written = AesCipher.encrypt(cipher, input, output, key, iv);

        assertEquals(16, written); // Padded to 16 bytes
        output.flip();

        // Verify by decrypting
        byte[] encrypted = new byte[written];
        output.get(encrypted);
        byte[] decrypted = AesCipher.decrypt(cipher, encrypted, key, iv);
        assertArrayEquals(plaintext, Arrays.copyOf(decrypted, plaintext.length));
    }

    @Test
    void encryptOffsetLengthDirect() throws GeneralSecurityException {
        byte[] key = new byte[16];
        byte[] iv = new byte[16];
        Arrays.fill(key, (byte) 0x42);
        Arrays.fill(iv, (byte) 0x24);

        // Data with prefix/suffix
        byte[] fullData = "XXXHelloYYY".getBytes(StandardCharsets.UTF_8);
        int offset = 3;
        int length = 5; // "Hello"

        byte[] ciphertext = AesCipher.encrypt(fullData, offset, length, key, iv);

        // Should be padded to 16 bytes
        assertEquals(16, ciphertext.length);

        // Decrypt and verify
        byte[] decrypted = AesCipher.decrypt(ciphertext, key, iv);
        byte[] expected = "Hello".getBytes(StandardCharsets.UTF_8);
        assertArrayEquals(expected, Arrays.copyOf(decrypted, expected.length));
    }

    @Test
    void allKeySizesProduceDifferentCiphertext() throws GeneralSecurityException {
        byte[] iv = new byte[16];
        byte[] plaintext = "Test message 16!".getBytes(StandardCharsets.UTF_8);

        byte[] key128 = new byte[16];
        byte[] key192 = new byte[24];
        byte[] key256 = new byte[32];
        Arrays.fill(key128, (byte) 0x42);
        Arrays.fill(key192, (byte) 0x42);
        Arrays.fill(key256, (byte) 0x42);

        byte[] ciphertext128 = AesCipher.encrypt(plaintext, key128, iv);
        byte[] ciphertext192 = AesCipher.encrypt(plaintext, key192, iv);
        byte[] ciphertext256 = AesCipher.encrypt(plaintext, key256, iv);

        // All should be different
        assertFalse(Arrays.equals(ciphertext128, ciphertext192));
        assertFalse(Arrays.equals(ciphertext128, ciphertext256));
        assertFalse(Arrays.equals(ciphertext192, ciphertext256));
    }
}
