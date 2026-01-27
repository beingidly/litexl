package com.beingidly.litexl.crypto;

import java.io.*;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.Channels;
import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import javax.crypto.Cipher;

/**
 * Writer for CFB (Compound File Binary) format used by encrypted Office documents.
 * Based on MS-CFB and MS-OFFCRYPTO specifications.
 *
 * <p>This class implements {@link Closeable} and should be used with try-with-resources:</p>
 * <pre>{@code
 * try (CfbWriter writer = new CfbWriter(out, encryptionInfo, encryptionKey, keyDataSalt)) {
 *     writer.writeEncrypted(plainDataStream, plainDataSize);
 * }
 * }</pre>
 */
public final class CfbWriter implements Closeable {

    // Reusable zero buffer for padding operations
    private static final byte[] ZERO_BUFFER = new byte[4096];

    // CFB constants (v3 format with 512-byte sectors)
    private static final int HEADER_SIZE = 512;
    private static final int SECTOR_SIZE = 512;
    private static final int MINI_SECTOR_SIZE = 64;
    private static final int MINI_STREAM_CUTOFF = 4096;
    private static final int DIR_ENTRY_SIZE = 128;
    private static final int DIFAT_IN_HEADER = 109;

    // Special sector values
    private static final int ENDOFCHAIN = 0xFFFFFFFE;
    private static final int FREESECT = 0xFFFFFFFF;
    private static final int FATSECT = 0xFFFFFFFD;

    // Directory entry types
    private static final byte OBJ_UNKNOWN = 0;
    private static final byte OBJ_STORAGE = 1;
    private static final byte OBJ_STREAM = 2;
    private static final byte OBJ_ROOT = 5;

    // DataSpaces static content (from MS-OFFCRYPTO spec)

    /** \x06DataSpaces/Version stream content */
    private static final byte[] DATASPACES_VERSION = {
        // FeatureIdentifier length (60 bytes = 30 chars)
        0x3C, 0x00, 0x00, 0x00,
        // "Microsoft.Container.DataSpaces" in UTF-16LE
        0x4D, 0x00, 0x69, 0x00, 0x63, 0x00, 0x72, 0x00, 0x6F, 0x00, 0x73, 0x00,
        0x6F, 0x00, 0x66, 0x00, 0x74, 0x00, 0x2E, 0x00, 0x43, 0x00, 0x6F, 0x00,
        0x6E, 0x00, 0x74, 0x00, 0x61, 0x00, 0x69, 0x00, 0x6E, 0x00, 0x65, 0x00,
        0x72, 0x00, 0x2E, 0x00, 0x44, 0x00, 0x61, 0x00, 0x74, 0x00, 0x61, 0x00,
        0x53, 0x00, 0x70, 0x00, 0x61, 0x00, 0x63, 0x00, 0x65, 0x00, 0x73, 0x00,
        // Reader version (1.0)
        0x01, 0x00, 0x00, 0x00,
        // Updater version (1.0)
        0x01, 0x00, 0x00, 0x00,
        // Writer version (1.0)
        0x01, 0x00, 0x00, 0x00
    };

    /** \x06DataSpaces/DataSpaceMap stream content */
    private static final byte[] DATASPACES_MAP = {
        // Header length
        0x08, 0x00, 0x00, 0x00,
        // Entry count
        0x01, 0x00, 0x00, 0x00,
        // Entry 1: MapEntry length
        0x68, 0x00, 0x00, 0x00,
        // Reference count
        0x01, 0x00, 0x00, 0x00,
        // Reference 1: ReferenceComponentType (0 = stream)
        0x00, 0x00, 0x00, 0x00,
        // Reference 1: Name length (32 bytes = "EncryptedPackage" in UTF-16LE)
        0x20, 0x00, 0x00, 0x00,
        // "EncryptedPackage" in UTF-16LE
        0x45, 0x00, 0x6E, 0x00, 0x63, 0x00, 0x72, 0x00, 0x79, 0x00, 0x70, 0x00,
        0x74, 0x00, 0x65, 0x00, 0x64, 0x00, 0x50, 0x00, 0x61, 0x00, 0x63, 0x00,
        0x6B, 0x00, 0x61, 0x00, 0x67, 0x00, 0x65, 0x00,
        // DataSpace name length (50 bytes = "StrongEncryptionDataSpace" in UTF-16LE)
        0x32, 0x00, 0x00, 0x00,
        // "StrongEncryptionDataSpace" in UTF-16LE
        0x53, 0x00, 0x74, 0x00, 0x72, 0x00, 0x6F, 0x00, 0x6E, 0x00, 0x67, 0x00,
        0x45, 0x00, 0x6E, 0x00, 0x63, 0x00, 0x72, 0x00, 0x79, 0x00, 0x70, 0x00,
        0x74, 0x00, 0x69, 0x00, 0x6F, 0x00, 0x6E, 0x00, 0x44, 0x00, 0x61, 0x00,
        0x74, 0x00, 0x61, 0x00, 0x53, 0x00, 0x70, 0x00, 0x61, 0x00, 0x63, 0x00,
        0x65, 0x00,
        // Padding
        0x00, 0x00
    };

    /** \x06DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace stream content */
    private static final byte[] DATASPACES_INFO = {
        // Header length
        0x08, 0x00, 0x00, 0x00,
        // TransformReference count
        0x01, 0x00, 0x00, 0x00,
        // Transform name length (50 bytes = "StrongEncryptionTransform" in UTF-16LE)
        0x32, 0x00, 0x00, 0x00,
        // "StrongEncryptionTransform" in UTF-16LE
        0x53, 0x00, 0x74, 0x00, 0x72, 0x00, 0x6F, 0x00, 0x6E, 0x00, 0x67, 0x00,
        0x45, 0x00, 0x6E, 0x00, 0x63, 0x00, 0x72, 0x00, 0x79, 0x00, 0x70, 0x00,
        0x74, 0x00, 0x69, 0x00, 0x6F, 0x00, 0x6E, 0x00, 0x54, 0x00, 0x72, 0x00,
        0x61, 0x00, 0x6E, 0x00, 0x73, 0x00, 0x66, 0x00, 0x6F, 0x00, 0x72, 0x00,
        0x6D, 0x00,
        // Padding
        0x00, 0x00
    };

    /** \x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary stream content */
    private static final byte[] DATASPACES_PRIMARY = {
        // TransformInfoHeader length
        0x58, 0x00, 0x00, 0x00,
        // TransformType (1 = Password transform)
        0x01, 0x00, 0x00, 0x00,
        // TransformID length (76 bytes = GUID string in UTF-16LE)
        0x4C, 0x00, 0x00, 0x00,
        // "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}" in UTF-16LE
        0x7B, 0x00, 0x46, 0x00, 0x46, 0x00, 0x39, 0x00, 0x41, 0x00, 0x33, 0x00,
        0x46, 0x00, 0x30, 0x00, 0x33, 0x00, 0x2D, 0x00, 0x35, 0x00, 0x36, 0x00,
        0x45, 0x00, 0x46, 0x00, 0x2D, 0x00, 0x34, 0x00, 0x36, 0x00, 0x31, 0x00,
        0x33, 0x00, 0x2D, 0x00, 0x42, 0x00, 0x44, 0x00, 0x44, 0x00, 0x35, 0x00,
        0x2D, 0x00, 0x35, 0x00, 0x41, 0x00, 0x34, 0x00, 0x31, 0x00, 0x43, 0x00,
        0x31, 0x00, 0x44, 0x00, 0x30, 0x00, 0x37, 0x00, 0x32, 0x00, 0x34, 0x00,
        0x36, 0x00, 0x7D, 0x00,
        // TransformName length (78 bytes)
        0x4E, 0x00, 0x00, 0x00,
        // "Microsoft.Container.EncryptionTransform" in UTF-16LE
        0x4D, 0x00, 0x69, 0x00, 0x63, 0x00, 0x72, 0x00, 0x6F, 0x00, 0x73, 0x00,
        0x6F, 0x00, 0x66, 0x00, 0x74, 0x00, 0x2E, 0x00, 0x43, 0x00, 0x6F, 0x00,
        0x6E, 0x00, 0x74, 0x00, 0x61, 0x00, 0x69, 0x00, 0x6E, 0x00, 0x65, 0x00,
        0x72, 0x00, 0x2E, 0x00, 0x45, 0x00, 0x6E, 0x00, 0x63, 0x00, 0x72, 0x00,
        0x79, 0x00, 0x70, 0x00, 0x74, 0x00, 0x69, 0x00, 0x6F, 0x00, 0x6E, 0x00,
        0x54, 0x00, 0x72, 0x00, 0x61, 0x00, 0x6E, 0x00, 0x73, 0x00, 0x66, 0x00,
        0x6F, 0x00, 0x72, 0x00, 0x6D, 0x00,
        // Null terminator
        0x00, 0x00,
        // Reader version (1.0)
        0x01, 0x00, 0x00, 0x00,
        // Updater version (1.0)
        0x01, 0x00, 0x00, 0x00,
        // Writer version (1.0)
        0x01, 0x00, 0x00, 0x00,
        // ExtensibilityHeader
        0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00,
        0x04, 0x00, 0x00, 0x00   // ExtensibilityVersion = 4
    };

    /**
     * Directory structure indices for encrypted file:
     *  0: Root Entry (storage)
     *  1: \x06DataSpaces (storage)
     *  2: EncryptedPackage (stream)
     *  3: EncryptionInfo (stream)
     *  4: DataSpaceInfo (storage)
     *  5: DataSpaceMap (stream)
     *  6: TransformInfo (storage)
     *  7: Version (stream)
     *  8: StrongEncryptionDataSpace (stream)
     *  9: StrongEncryptionTransform (storage)
     * 10: \x06Primary (stream)
     */
    private static final int DIR_COUNT = 11;

    private static final int SEGMENT_SIZE = 4096;
    private static final int BLOCK_SIZE = 16;
    private static final int MAX_PADDED_LEN = ((SEGMENT_SIZE + BLOCK_SIZE - 1) / BLOCK_SIZE) * BLOCK_SIZE;

    private final OutputStream out;
    private final byte[] encryptionInfo;
    private final byte[] encryptionKey;
    private final byte[] keyDataSalt;

    // Reusable buffers (allocated once per writer instance, released in close())
    private final java.nio.channels.WritableByteChannel outChannel;
    private final ByteBuffer sizeBuf;
    private final ByteBuffer segmentInputBuf;
    private final ByteBuffer encryptOutputBuf;
    private final ByteBuffer indexBuf;
    private final byte[] ivBuffer = new byte[16];
    private final byte[] readBuffer = new byte[SEGMENT_SIZE];
    private final Cipher cipher;

    /**
     * Creates a new CFB writer.
     *
     * @param out the output stream to write to
     * @param encryptionInfo the encryption info XML with header
     * @param encryptionKey the AES encryption key
     * @param keyDataSalt the salt for IV generation
     */
    public CfbWriter(OutputStream out, byte[] encryptionInfo, byte[] encryptionKey, byte[] keyDataSalt) {
        this.out = out;
        this.encryptionInfo = encryptionInfo;
        this.encryptionKey = encryptionKey;
        this.keyDataSalt = keyDataSalt;
        this.outChannel = Channels.newChannel(out);

        // Allocate direct buffers (reused across operations, released in close())
        this.sizeBuf = ByteBuffer.allocateDirect(8).order(ByteOrder.LITTLE_ENDIAN);
        this.segmentInputBuf = ByteBuffer.allocateDirect(SEGMENT_SIZE);
        this.encryptOutputBuf = ByteBuffer.allocateDirect(MAX_PADDED_LEN);
        this.indexBuf = ByteBuffer.allocateDirect(4).order(ByteOrder.LITTLE_ENDIAN);
        try {
            this.cipher = Cipher.getInstance("AES/CBC/NoPadding");
        } catch (GeneralSecurityException e) {
            throw new RuntimeException("AES cipher not available", e);
        }
    }

    /**
     * Writes an encrypted Office document with streaming encryption.
     * Memory usage is O(1) regardless of file size.
     *
     * @param plainDataStream input stream providing plain data
     * @param plainDataSize size of the plain data to encrypt
     * @throws IOException if an I/O error occurs
     * @throws GeneralSecurityException if encryption fails
     */
    public void writeEncrypted(InputStream plainDataStream, long plainDataSize)
            throws IOException, GeneralSecurityException {

        long encryptedPackageSize = calculateEncryptedPackageSize(plainDataSize);

        // Build small streams array (EncryptedPackage is streamed separately)
        byte[][] streams = new byte[DIR_COUNT][];
        streams[3] = encryptionInfo;    // EncryptionInfo
        streams[5] = DATASPACES_MAP;    // DataSpaceMap
        streams[7] = DATASPACES_VERSION; // Version
        streams[8] = DATASPACES_INFO;   // StrongEncryptionDataSpace
        streams[10] = DATASPACES_PRIMARY; // \x06Primary

        // Stream sizes
        long[] streamSizes = new long[DIR_COUNT];
        streamSizes[2] = encryptedPackageSize;
        for (int i = 0; i < DIR_COUNT; i++) {
            if (i != 2 && streams[i] != null) {
                streamSizes[i] = streams[i].length;
            }
        }

        // Calculate mini stream size (streams < 4096 bytes, except EncryptedPackage)
        int miniStreamTotal = 0;
        for (int i = 0; i < DIR_COUNT; i++) {
            if (i == 2) {
                continue;
            }
            if (streamSizes[i] > 0 && streamSizes[i] < MINI_STREAM_CUTOFF) {
                miniStreamTotal += sectorsNeeded((int) streamSizes[i], MINI_SECTOR_SIZE) * MINI_SECTOR_SIZE;
            }
        }

        // Calculate sector counts
        int encryptedPackageSectors = sectorsNeeded(encryptedPackageSize, SECTOR_SIZE);

        int miniStreamSectors = sectorsNeeded(miniStreamTotal, SECTOR_SIZE);
        int miniSectorCount = miniStreamTotal / MINI_SECTOR_SIZE;
        int miniFatSectors = sectorsNeeded(miniSectorCount * 4, SECTOR_SIZE);
        int dirSectors = sectorsNeeded(DIR_COUNT * DIR_ENTRY_SIZE, SECTOR_SIZE);

        // Calculate FAT size iteratively
        int fatSectors = 1;
        int totalSectors;
        do {
            totalSectors = fatSectors + miniFatSectors + dirSectors + miniStreamSectors + encryptedPackageSectors;
            int newFat = sectorsNeeded((long) totalSectors * 4, SECTOR_SIZE);
            if (newFat == fatSectors) {
                break;
            }
            fatSectors = newFat;
        } while (fatSectors < 1000);

        // Sector layout
        int firstMiniFat = fatSectors;
        int firstDir = firstMiniFat + miniFatSectors;
        int firstMiniStream = firstDir + dirSectors;

        // Metadata buffer: header + FAT + miniFAT + directory + mini streams (Direct ByteBuffer)
        int metadataSize = HEADER_SIZE + (fatSectors + miniFatSectors + dirSectors + miniStreamSectors) * SECTOR_SIZE;
        ByteBuffer metadataBuffer = ByteBuffer.allocateDirect(metadataSize).order(ByteOrder.LITTLE_ENDIAN);

        // Write header
        writeHeader(metadataBuffer, fatSectors, firstDir, miniFatSectors, firstMiniFat);

        // Write FAT
        int[] regularStarts = new int[DIR_COUNT];
        writeFat(metadataBuffer, fatSectors, miniFatSectors, dirSectors,
                 miniStreamSectors, encryptedPackageSectors, regularStarts);

        // Write Mini FAT and Mini Stream
        int[] miniStarts = new int[DIR_COUNT];
        writeMiniStream(metadataBuffer, firstMiniFat, firstMiniStream, miniFatSectors,
                        streams, streamSizes, miniStarts);

        // Write directory entries
        writeDirectory(metadataBuffer, firstDir, dirSectors, streamSizes,
                       miniStarts, regularStarts, firstMiniStream, miniStreamTotal);

        // Write metadata to output via channel
        metadataBuffer.position(0);
        metadataBuffer.limit(metadataSize);
        outChannel.write(metadataBuffer);

        // Stream the encrypted package
        streamEncryptedData(plainDataStream, plainDataSize, encryptedPackageSectors);
    }

    /**
     * Calculates the encrypted package size for a given plain data size.
     */
    private long calculateEncryptedPackageSize(long plainDataSize) {
        int numSegments = (int) ((plainDataSize + SEGMENT_SIZE - 1) / SEGMENT_SIZE);
        long encryptedSize = 8; // 8-byte size header
        for (int i = 0; i < numSegments; i++) {
            int segmentLen = (int) Math.min(SEGMENT_SIZE, plainDataSize - (long) i * SEGMENT_SIZE);
            int paddedLen = ((segmentLen + BLOCK_SIZE - 1) / BLOCK_SIZE) * BLOCK_SIZE;
            encryptedSize += paddedLen;
        }
        // Ensure minimum size of 4104 bytes for regular sectors (per MS-OFFCRYPTO spec)
        return Math.max(encryptedSize, 4104);
    }

    /**
     * Streams plain data, encrypting in 4KB segments and writing directly to output.
     */
    private void streamEncryptedData(InputStream plainData, long plainDataSize,
            int encryptedPackageSectors)
            throws IOException, GeneralSecurityException {

        // Write 8-byte size header (original plain data size)
        sizeBuf.clear();
        sizeBuf.putLong(plainDataSize);
        sizeBuf.flip();
        outChannel.write(sizeBuf);

        MessageDigest sha512 = MessageDigest.getInstance("SHA-512");

        long bytesWritten = 8; // size header already written
        int segmentIndex = 0;
        long remaining = plainDataSize;

        while (remaining > 0) {
            int toRead = (int) Math.min(SEGMENT_SIZE, remaining);
            int totalRead = 0;

            // Read full segment (may require multiple reads)
            while (totalRead < toRead) {
                int read = plainData.read(readBuffer, totalRead, toRead - totalRead);
                if (read == -1) {
                    break;
                }
                totalRead += read;
            }

            if (totalRead == 0) {
                break;
            }

            // Copy to direct buffer for encryption
            segmentInputBuf.clear();
            segmentInputBuf.put(readBuffer, 0, totalRead);
            segmentInputBuf.flip();

            // Generate IV: SHA512(salt + LE32(segmentIndex))[0:16]
            sha512.reset();
            sha512.update(keyDataSalt);
            indexBuf.clear();
            indexBuf.putInt(segmentIndex);
            indexBuf.flip();
            sha512.update(indexBuf);
            byte[] hash = sha512.digest();
            System.arraycopy(hash, 0, ivBuffer, 0, 16);

            // Encrypt segment using ByteBuffer API with reusable Cipher
            encryptOutputBuf.clear();
            int encryptedLen = AesCipher.encrypt(cipher, segmentInputBuf, encryptOutputBuf, encryptionKey, ivBuffer);
            encryptOutputBuf.flip();
            outChannel.write(encryptOutputBuf);
            bytesWritten += encryptedLen;

            remaining -= totalRead;
            segmentIndex++;
        }

        // Pad to minimum size if needed (4104 bytes minimum)
        int paddingRemaining = (int) (4104 - bytesWritten);
        while (paddingRemaining > 0) {
            int toWrite = Math.min(paddingRemaining, ZERO_BUFFER.length);
            out.write(ZERO_BUFFER, 0, toWrite);
            paddingRemaining -= toWrite;
            bytesWritten += toWrite;
        }

        // Pad to sector boundary
        int sectorPaddingRemaining = (int) ((long) encryptedPackageSectors * SECTOR_SIZE - bytesWritten);
        while (sectorPaddingRemaining > 0) {
            int toWrite = Math.min(sectorPaddingRemaining, ZERO_BUFFER.length);
            out.write(ZERO_BUFFER, 0, toWrite);
            sectorPaddingRemaining -= toWrite;
        }
    }

    private int sectorsNeeded(long size, int sectorSize) {
        if (size == 0) {
            return 0;
        }
        return (int) ((size + sectorSize - 1) / sectorSize);
    }

    private void writeFat(ByteBuffer buffer, int fatSectors, int miniFatSectors,
            int dirSectors, int miniStreamSectors, int encryptedPackageSectors, int[] regularStarts) {
        buffer.position(HEADER_SIZE);
        int sector = 0;

        // FAT sectors
        for (int i = 0; i < fatSectors; i++) {
            buffer.putInt(FATSECT);
            sector++;
        }

        // Mini FAT sectors chain
        for (int i = 0; i < miniFatSectors; i++) {
            buffer.putInt(i < miniFatSectors - 1 ? sector + 1 : ENDOFCHAIN);
            sector++;
        }

        // Directory sectors chain
        for (int i = 0; i < dirSectors; i++) {
            buffer.putInt(i < dirSectors - 1 ? sector + 1 : ENDOFCHAIN);
            sector++;
        }

        // Mini stream sectors chain
        for (int i = 0; i < miniStreamSectors; i++) {
            buffer.putInt(i < miniStreamSectors - 1 ? sector + 1 : ENDOFCHAIN);
            sector++;
        }

        // EncryptedPackage sectors (index 2)
        regularStarts[2] = sector;
        for (int i = 0; i < encryptedPackageSectors; i++) {
            buffer.putInt(i < encryptedPackageSectors - 1 ? sector + 1 : ENDOFCHAIN);
            sector++;
        }

        // Fill remaining FAT with FREESECT
        int fatEntries = fatSectors * (SECTOR_SIZE / 4);
        for (int i = sector; i < fatEntries; i++) {
            buffer.putInt(FREESECT);
        }
    }

    private void writeMiniStream(ByteBuffer buffer, int firstMiniFat, int firstMiniStream,
            int miniFatSectors, byte[][] streams, long[] streamSizes, int[] miniStarts) {
        if (miniFatSectors == 0) {
            return;
        }

        int miniFatOffset = HEADER_SIZE + firstMiniFat * SECTOR_SIZE;
        int miniStreamOffset = HEADER_SIZE + firstMiniStream * SECTOR_SIZE;
        int miniSector = 0;
        int dataOffset = 0;

        for (int i = 0; i < streams.length; i++) {
            if (i == 2) {
                continue; // EncryptedPackage goes to regular sectors
            }
            if (streams[i] != null && streamSizes[i] > 0 && streamSizes[i] < MINI_STREAM_CUTOFF) {
                miniStarts[i] = miniSector;
                int mSectors = sectorsNeeded((int) streamSizes[i], MINI_SECTOR_SIZE);

                // Copy data to mini stream
                buffer.position(miniStreamOffset + dataOffset);
                buffer.put(streams[i]);

                // Build mini FAT chain
                buffer.position(miniFatOffset + miniSector * 4);
                for (int j = 0; j < mSectors; j++) {
                    buffer.putInt(j < mSectors - 1 ? miniSector + 1 : ENDOFCHAIN);
                    miniSector++;
                }

                dataOffset += mSectors * MINI_SECTOR_SIZE;
            }
        }

        // Fill remaining Mini FAT with FREESECT
        int miniFatEntries = miniFatSectors * (SECTOR_SIZE / 4);
        buffer.position(miniFatOffset + miniSector * 4);
        for (int i = miniSector; i < miniFatEntries; i++) {
            buffer.putInt(FREESECT);
        }
    }

    private void writeDirectory(ByteBuffer buffer, int firstDir, int dirSectors,
            long[] streamSizes, int[] miniStarts, int[] regularStarts,
            int firstMiniStream, int miniStreamTotal) {
        int dirOffset = HEADER_SIZE + firstDir * SECTOR_SIZE;

        // Initialize all directory entries
        for (int i = 0; i < dirSectors * (SECTOR_SIZE / DIR_ENTRY_SIZE); i++) {
            int entryOffset = dirOffset + i * DIR_ENTRY_SIZE;
            buffer.position(entryOffset + 66);
            buffer.put(OBJ_UNKNOWN);
            buffer.put((byte) 1); // color = black
            buffer.putInt(-1);
            buffer.putInt(-1);
            buffer.putInt(-1);
        }

        // Entry 0: Root Entry
        writeEntry(buffer, dirOffset, 0, "Root Entry", false, OBJ_ROOT,
                   -1, -1, 3,
                   miniStreamTotal > 0 ? firstMiniStream : ENDOFCHAIN, miniStreamTotal);

        // Entry 1: \x06DataSpaces (storage)
        writeEntry(buffer, dirOffset, 1, "DataSpaces", true, OBJ_STORAGE,
                   -1, -1, 5, 0, 0);

        // Entry 2: EncryptedPackage (stream)
        writeEntry(buffer, dirOffset, 2, "EncryptedPackage", false, OBJ_STREAM,
                   -1, -1, -1, regularStarts[2], streamSizes[2]);

        // Entry 3: EncryptionInfo (stream)
        int encInfoStart = streamSizes[3] >= MINI_STREAM_CUTOFF ? regularStarts[3] : miniStarts[3];
        writeEntry(buffer, dirOffset, 3, "EncryptionInfo", false, OBJ_STREAM,
                   1, 2, -1, encInfoStart, streamSizes[3]);

        // Entry 4: DataSpaceInfo (storage)
        writeEntry(buffer, dirOffset, 4, "DataSpaceInfo", false, OBJ_STORAGE,
                   -1, -1, 8, 0, 0);

        // Entry 5: DataSpaceMap (stream)
        writeEntry(buffer, dirOffset, 5, "DataSpaceMap", false, OBJ_STREAM,
                   4, 6, -1, miniStarts[5], streamSizes[5]);

        // Entry 6: TransformInfo (storage)
        writeEntry(buffer, dirOffset, 6, "TransformInfo", false, OBJ_STORAGE,
                   -1, 7, 9, 0, 0);

        // Entry 7: Version (stream)
        writeEntry(buffer, dirOffset, 7, "Version", false, OBJ_STREAM,
                   -1, -1, -1, miniStarts[7], streamSizes[7]);

        // Entry 8: StrongEncryptionDataSpace (stream)
        writeEntry(buffer, dirOffset, 8, "StrongEncryptionDataSpace", false, OBJ_STREAM,
                   -1, -1, -1, miniStarts[8], streamSizes[8]);

        // Entry 9: StrongEncryptionTransform (storage)
        writeEntry(buffer, dirOffset, 9, "StrongEncryptionTransform", false, OBJ_STORAGE,
                   -1, -1, 10, 0, 0);

        // Entry 10: \x06Primary (stream)
        writeEntry(buffer, dirOffset, 10, "Primary", true, OBJ_STREAM,
                   -1, -1, -1, miniStarts[10], streamSizes[10]);
    }

    private void writeHeader(ByteBuffer buffer, int fatSectors, int firstDir,
                             int miniFatSectors, int firstMiniFat) {
        buffer.position(0);

        // CFB Signature: D0 CF 11 E0 A1 B1 1A E1
        buffer.put((byte) 0xD0);
        buffer.put((byte) 0xCF);
        buffer.put((byte) 0x11);
        buffer.put((byte) 0xE0);
        buffer.put((byte) 0xA1);
        buffer.put((byte) 0xB1);
        buffer.put((byte) 0x1A);
        buffer.put((byte) 0xE1);

        // CLSID (16 bytes of zeros)
        buffer.position(0x08);
        buffer.putLong(0);
        buffer.putLong(0);

        // Version
        buffer.position(0x18);
        buffer.putShort((short) 0x003E);  // Minor version
        buffer.putShort((short) 0x0003);  // Major version (v3)
        buffer.putShort((short) 0xFFFE);  // Byte order (little-endian)
        buffer.putShort((short) 0x0009);  // Sector size power (512)
        buffer.putShort((short) 0x0006);  // Mini sector size power (64)

        buffer.position(0x2C);
        buffer.putInt(fatSectors);        // Number of FAT sectors
        buffer.putInt(firstDir);          // First directory sector
        buffer.putInt(0);                 // Transaction signature
        buffer.putInt(MINI_STREAM_CUTOFF); // Mini stream cutoff

        if (miniFatSectors > 0) {
            buffer.putInt(firstMiniFat);  // First mini FAT sector
            buffer.putInt(miniFatSectors); // Number of mini FAT sectors
        } else {
            buffer.putInt(ENDOFCHAIN);
            buffer.putInt(0);
        }

        buffer.putInt(ENDOFCHAIN);        // First DIFAT sector
        buffer.putInt(0);                 // Number of DIFAT sectors

        // DIFAT array (109 entries in header)
        for (int i = 0; i < DIFAT_IN_HEADER; i++) {
            buffer.putInt(i < fatSectors ? i : FREESECT);
        }
    }

    private void writeEntry(ByteBuffer dirBuf, int dirOffset, int entryIndex, String name, boolean hasPrefix06,
                            byte type, int leftSibling, int rightSibling, int child,
                            int startSector, long size) {
        int offset = dirOffset + entryIndex * DIR_ENTRY_SIZE;

        // Write name in UTF-16LE
        dirBuf.position(offset);
        int nameLen = 0;
        if (hasPrefix06) {
            dirBuf.putShort((short) 0x06);
            nameLen += 2;
        }
        for (int i = 0; i < name.length() && nameLen < 62; i++) {
            dirBuf.putShort((short) name.charAt(i));
            nameLen += 2;
        }
        dirBuf.putShort((short) 0); // Null terminator
        nameLen += 2;

        dirBuf.position(offset + 64);
        dirBuf.putShort((short) nameLen); // Name length
        dirBuf.put(type);
        dirBuf.put((byte) 1); // color = black
        dirBuf.putInt(leftSibling);
        dirBuf.putInt(rightSibling);
        dirBuf.putInt(child);
        // CLSID (16 bytes) at offset 80 - already zeros
        // State bits at offset 96 - already zeros
        // Timestamps at offset 100, 108 - already zeros
        dirBuf.position(offset + 116);
        dirBuf.putInt(startSector);
        dirBuf.putLong(size);
    }

    @Override
    public void close() throws IOException {
        out.close();
    }
}
