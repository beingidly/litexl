package com.beingidly.litexl;

import com.beingidly.litexl.crypto.AgileDecryptor;
import org.junit.jupiter.api.Test;
import java.nio.ByteBuffer;
import java.nio.file.Path;
import static org.junit.jupiter.api.Assertions.*;

class AgileDecryptorTest {
    private static final String TEST_PASSWORD = "password123";

    @Test
    void decryptsAes256EncryptedFile() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes256.xlsx");
        try (var cfbReader = new CfbReader(encrypted)) {
            var decryptor = new AgileDecryptor(TEST_PASSWORD);
            decryptor.parseEncryptionInfo(
                cfbReader.getEncryptionInfoBuffer(),
                cfbReader.getEncryptedPackageBuffer()
            );
            ByteBuffer decrypted = decryptor.decrypt();
            assertNotNull(decrypted);
            assertTrue(decrypted.remaining() > 0);
            assertEquals(0x50, decrypted.get(0) & 0xFF); // 'P'
            assertEquals(0x4B, decrypted.get(1) & 0xFF); // 'K'
        }
    }

    @Test
    void decryptsAes128EncryptedFile() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes128.xlsx");
        try (var cfbReader = new CfbReader(encrypted)) {
            var decryptor = new AgileDecryptor(TEST_PASSWORD);
            decryptor.parseEncryptionInfo(
                cfbReader.getEncryptionInfoBuffer(),
                cfbReader.getEncryptedPackageBuffer()
            );
            ByteBuffer decrypted = decryptor.decrypt();
            assertNotNull(decrypted);
            assertEquals(0x50, decrypted.get(0) & 0xFF);
            assertEquals(0x4B, decrypted.get(1) & 0xFF);
        }
    }

    @Test
    void throwsOnWrongPassword() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes256.xlsx");
        try (var cfbReader = new CfbReader(encrypted)) {
            var decryptor = new AgileDecryptor("wrongpassword");
            decryptor.parseEncryptionInfo(
                cfbReader.getEncryptionInfoBuffer(),
                cfbReader.getEncryptedPackageBuffer()
            );
            assertThrows(InvalidPasswordException.class, decryptor::decrypt);
        }
    }

    @Test
    void throwsOnNullEncryptionInfo() {
        var decryptor = new AgileDecryptor(TEST_PASSWORD);
        assertThrows(NullPointerException.class, () ->
            decryptor.parseEncryptionInfo(null, ByteBuffer.allocate(10))
        );
    }

    @Test
    void throwsOnMalformedXml() {
        var decryptor = new AgileDecryptor(TEST_PASSWORD);
        ByteBuffer malformedInfo = ByteBuffer.wrap("<invalid>".getBytes());
        ByteBuffer encryptedPackage = ByteBuffer.allocate(100);
        assertThrows(Exception.class, () ->
            decryptor.parseEncryptionInfo(malformedInfo, encryptedPackage)
        );
    }
}
