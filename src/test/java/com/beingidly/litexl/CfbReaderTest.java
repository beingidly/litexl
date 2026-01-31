package com.beingidly.litexl;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

class CfbReaderTest {

    @TempDir
    Path tempDir;

    @Test
    void readsEncryptedExcelFile() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes256.xlsx");

        try (var reader = new CfbReader(encrypted)) {
            assertTrue(reader.isEncrypted());
            assertNotNull(reader.getEncryptionInfoBuffer());
            assertNotNull(reader.getEncryptedPackageBuffer());
        }
    }

    @Test
    void detectsNonEncryptedFile() throws Exception {
        // Create a valid CFB file without encryption entries
        Path nonEncryptedCfb = createMinimalCfb();

        try (var reader = new CfbReader(nonEncryptedCfb)) {
            assertFalse(reader.isEncrypted());
        }
    }

    @Test
    void throwsOnZipFile() throws Exception {
        // test-data.xlsx is a ZIP file (unencrypted XLSX), not a CFB file
        Path plain = Path.of("src/test/resources/test-data.xlsx");

        // CfbReader should throw because ZIP files have different signature
        assertThrows(CorruptFileException.class, () -> new CfbReader(plain));
    }

    @Test
    void throwsOnInvalidCfbSignature() throws Exception {
        Path invalid = tempDir.resolve("invalid.xlsx");
        Files.write(invalid, new byte[]{0x00, 0x01, 0x02, 0x03});

        assertThrows(CorruptFileException.class, () -> new CfbReader(invalid));
    }

    @Test
    void throwsOnInvalidByteOrder() throws Exception {
        Path invalid = tempDir.resolve("bad-order.cfb");
        byte[] data = new byte[512];
        // Valid CFB signature
        data[0] = (byte) 0xD0; data[1] = (byte) 0xCF;
        data[2] = (byte) 0x11; data[3] = (byte) 0xE0;
        data[4] = (byte) 0xA1; data[5] = (byte) 0xB1;
        data[6] = (byte) 0x1A; data[7] = (byte) 0xE1;
        // Invalid byte order at offset 0x1C (should be 0xFFFE for little-endian)
        data[0x1C] = (byte) 0x00;
        data[0x1D] = (byte) 0x00;
        Files.write(invalid, data);

        assertThrows(CorruptFileException.class, () -> new CfbReader(invalid));
    }

    @Test
    void returnsNullForMissingStreams() throws Exception {
        Path minimalCfb = createMinimalCfb();

        try (var reader = new CfbReader(minimalCfb)) {
            assertFalse(reader.isEncrypted());
            assertNull(reader.getEncryptionInfoBuffer());
            assertNull(reader.getEncryptedPackageBuffer());
        }
    }

    @Test
    void closesResources() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes256.xlsx");
        var reader = new CfbReader(encrypted);
        reader.close();
        reader.close(); // Should not throw on double close
    }

    @Test
    void readStreamAsBufferReturnsNullForNonExistentStream() throws Exception {
        Path minimalCfb = createMinimalCfb();

        try (var reader = new CfbReader(minimalCfb)) {
            // Test readStreamAsBuffer directly with a non-existent stream name
            assertNull(reader.readStreamAsBuffer("NonExistentStream"));
            assertNull(reader.readStreamAsBuffer("AnotherMissingStream"));
        }
    }

    @Test
    void readStreamAsBufferReturnsBufferForExistingStream() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes256.xlsx");

        try (var reader = new CfbReader(encrypted)) {
            // EncryptionInfo should exist in encrypted files
            ByteBuffer infoBuffer = reader.readStreamAsBuffer("EncryptionInfo");
            assertNotNull(infoBuffer);
            assertTrue(infoBuffer.remaining() > 0);

            // EncryptedPackage should also exist
            ByteBuffer packageBuffer = reader.readStreamAsBuffer("EncryptedPackage");
            assertNotNull(packageBuffer);
            assertTrue(packageBuffer.remaining() > 0);
        }
    }

    @Test
    void throwsOnTruncatedFile() throws Exception {
        // File too small to contain valid CFB header
        Path truncated = tempDir.resolve("truncated.cfb");
        byte[] data = new byte[]{
            (byte) 0xD0, (byte) 0xCF, (byte) 0x11, (byte) 0xE0,
            (byte) 0xA1, (byte) 0xB1, (byte) 0x1A, (byte) 0xE1
        };
        Files.write(truncated, data);

        // Should throw when trying to read header fields beyond file size
        assertThrows(Exception.class, () -> new CfbReader(truncated));
    }

    @Test
    void isEncryptedReturnsFalseWhenNoEncryptionEntries() throws Exception {
        Path minimalCfb = createMinimalCfb();

        try (var reader = new CfbReader(minimalCfb)) {
            // Call isEncrypted() multiple times to test caching behavior
            assertFalse(reader.isEncrypted());
            assertFalse(reader.isEncrypted());
        }
    }

    @Test
    void isEncryptedReturnsTrueForEncryptedFile() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes256.xlsx");

        try (var reader = new CfbReader(encrypted)) {
            // Call isEncrypted() multiple times to test caching behavior
            assertTrue(reader.isEncrypted());
            assertTrue(reader.isEncrypted());
        }
    }

    /**
     * Creates a minimal valid CFB file with a DummyStream entry (no encryption entries).
     * This is a simplified CFB v3 format with:
     * - 512-byte header
     * - 1 FAT sector
     * - 1 directory sector
     * - 1 mini stream sector (for DummyStream data)
     */
    private Path createMinimalCfb() throws Exception {
        Path path = tempDir.resolve("minimal.cfb");

        // CFB v3: 512-byte sectors
        int sectorSize = 512;
        int headerSize = 512;
        int fatSector = 0;
        int dirSector = 1;
        int miniStreamSector = 2;

        // Total file size: header + 3 sectors (FAT, directory, mini stream)
        int fileSize = headerSize + 3 * sectorSize;
        ByteBuffer buffer = ByteBuffer.allocate(fileSize).order(ByteOrder.LITTLE_ENDIAN);

        // Write header
        // CFB signature
        buffer.put((byte) 0xD0);
        buffer.put((byte) 0xCF);
        buffer.put((byte) 0x11);
        buffer.put((byte) 0xE0);
        buffer.put((byte) 0xA1);
        buffer.put((byte) 0xB1);
        buffer.put((byte) 0x1A);
        buffer.put((byte) 0xE1);

        // CLSID (16 bytes of zeros at offset 8)
        buffer.position(0x18);
        buffer.putShort((short) 0x003E);  // Minor version
        buffer.putShort((short) 0x0003);  // Major version (v3)
        buffer.putShort((short) 0xFFFE);  // Byte order (little-endian)
        buffer.putShort((short) 0x0009);  // Sector size power (2^9 = 512)
        buffer.putShort((short) 0x0006);  // Mini sector size power (2^6 = 64)

        buffer.position(0x2C);
        buffer.putInt(1);                 // Number of FAT sectors
        buffer.putInt(dirSector);         // First directory sector
        buffer.putInt(0);                 // Transaction signature
        buffer.putInt(4096);              // Mini stream cutoff

        buffer.putInt(miniStreamSector);  // First mini FAT sector (actually mini stream)
        buffer.putInt(0);                 // Number of mini FAT sectors
        buffer.putInt(-2);                // First DIFAT sector (ENDOFCHAIN)
        buffer.putInt(0);                 // Number of DIFAT sectors

        // DIFAT array (109 entries at offset 0x4C)
        buffer.position(0x4C);
        buffer.putInt(fatSector);         // First FAT sector
        for (int i = 1; i < 109; i++) {
            buffer.putInt(-1);            // FREESECT
        }

        // Write FAT sector (sector 0)
        buffer.position(headerSize + fatSector * sectorSize);
        buffer.putInt(0xFFFFFFFD);        // FAT[0] = FATSECT (this sector is a FAT)
        buffer.putInt(-2);                // FAT[1] = ENDOFCHAIN (directory sector)
        buffer.putInt(-2);                // FAT[2] = ENDOFCHAIN (mini stream sector)
        // Fill rest with FREESECT
        for (int i = 3; i < sectorSize / 4; i++) {
            buffer.putInt(-1);
        }

        // Write directory sector (sector 1)
        int dirOffset = headerSize + dirSector * sectorSize;

        // Entry 0: Root Entry (storage)
        writeDirectoryEntry(buffer, dirOffset, 0,
                "Root Entry", (byte) 5, // type = root
                -1, -1, 1,              // siblings, child
                miniStreamSector, 3);   // startSector, size (mini stream size)

        // Entry 1: DummyStream (stream)
        writeDirectoryEntry(buffer, dirOffset, 1,
                "DummyStream", (byte) 2, // type = stream
                -1, -1, -1,             // siblings, child
                0, 3);                  // startSector (mini sector 0), size

        // Fill remaining entries as unused
        for (int i = 2; i < sectorSize / 128; i++) {
            buffer.position(dirOffset + i * 128 + 66);
            buffer.put((byte) 0);       // type = unused
        }

        // Write mini stream sector (sector 2) - contains DummyStream data
        buffer.position(headerSize + miniStreamSector * sectorSize);
        buffer.put((byte) 1);
        buffer.put((byte) 2);
        buffer.put((byte) 3);

        Files.write(path, buffer.array());
        return path;
    }

    private void writeDirectoryEntry(ByteBuffer buffer, int dirOffset, int entryIndex,
                                     String name, byte type,
                                     int leftSibling, int rightSibling, int child,
                                     int startSector, long size) {
        int offset = dirOffset + entryIndex * 128;
        buffer.position(offset);

        // Write name in UTF-16LE
        int nameLen = 0;
        for (int i = 0; i < name.length() && nameLen < 62; i++) {
            buffer.putShort((short) name.charAt(i));
            nameLen += 2;
        }
        buffer.putShort((short) 0); // Null terminator
        nameLen += 2;

        buffer.position(offset + 64);
        buffer.putShort((short) nameLen); // Name length in bytes
        buffer.put(type);                 // Object type
        buffer.put((byte) 1);             // Color = black
        buffer.putInt(leftSibling);
        buffer.putInt(rightSibling);
        buffer.putInt(child);

        // CLSID (16 bytes at offset 80) - leave as zeros
        // State bits (4 bytes at offset 96) - leave as zeros
        // Created/modified times (16 bytes at offset 100) - leave as zeros

        buffer.position(offset + 116);
        buffer.putInt(startSector);
        buffer.putLong(size);
    }
}
