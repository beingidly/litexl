package com.beingidly.litexl;

import org.jspecify.annotations.Nullable;

import java.io.*;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.MappedByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.*;

/**
 * Reader for CFB (Compound File Binary) format used by encrypted Office documents.
 * Also known as OLE2 or Structured Storage format.
 *
 * <p>Uses memory-mapped I/O for zero-copy reads:
 * <ul>
 *   <li>No byte[] allocations for reading sectors</li>
 *   <li>FAT/MiniFAT entries read on-demand from mapped buffer</li>
 *   <li>Stream data returned as ByteBuffer slices when possible</li>
 * </ul>
 */
final class CfbReader implements Closeable {

    // Note: MS spec says B2 but POI uses B1
    private static final byte[] CFB_SIGNATURE = {
        (byte) 0xD0, (byte) 0xCF, (byte) 0x11, (byte) 0xE0,
        (byte) 0xA1, (byte) 0xB1, (byte) 0x1A, (byte) 0xE1
    };

    private static final int HEADER_SIZE = 512;
    private static final int DIR_ENTRY_SIZE = 128;
    private static final int ENDOFCHAIN = -2;

    private final FileChannel channel;
    private final MappedByteBuffer mappedBuffer;
    private final int sectorSize;
    private final int miniSectorSize;
    private final int miniStreamCutoffSize;
    private final int fatEntriesPerSector;

    // DIFAT metadata (no arrays - read on demand from mapped buffer)
    private final int numFatSectors;
    private final int firstDifatSector;

    // MiniFAT location
    private final int firstMiniFatSector;
    private final int numMiniFatSectors;

    // Directory: store sector chain start only, parse entries on-demand
    private final int firstDirSector;
    private final int miniStreamStart;

    // Cache for looked-up entries (lazy population)
    private volatile @Nullable DirectoryEntry encryptedPackageEntry;
    private volatile @Nullable DirectoryEntry encryptionInfoEntry;

    public CfbReader(Path path) throws IOException {
        // Memory-map the entire file for zero-copy reads
        this.channel = FileChannel.open(path, StandardOpenOption.READ);
        long fileSize = channel.size();
        this.mappedBuffer = channel.map(FileChannel.MapMode.READ_ONLY, 0, fileSize);
        mappedBuffer.order(ByteOrder.LITTLE_ENDIAN);

        // Verify signature (zero-copy)
        for (int i = 0; i < CFB_SIGNATURE.length; i++) {
            if (mappedBuffer.get(i) != CFB_SIGNATURE[i]) {
                throw new CorruptFileException("Invalid CFB signature");
            }
        }

        // Read byte order at offset 0x1C (28)
        if (mappedBuffer.getShort(28) != (short) 0xFFFE) {
            throw new CorruptFileException("Invalid byte order");
        }

        // Sector size power at offset 0x1E (30)
        this.sectorSize = 1 << mappedBuffer.getShort(30);
        this.fatEntriesPerSector = sectorSize / 4;

        // Mini sector size power at offset 0x20 (32)
        this.miniSectorSize = 1 << mappedBuffer.getShort(32);

        // Mini stream cutoff at offset 0x38 (56)
        this.miniStreamCutoffSize = mappedBuffer.getInt(56);

        // Store header values only - no arrays allocated
        // DIFAT entries read on-demand from header (offset 76) or DIFAT sectors
        this.numFatSectors = mappedBuffer.getInt(44);
        this.firstDirSector = mappedBuffer.getInt(48);
        this.firstMiniFatSector = mappedBuffer.getInt(60);
        this.numMiniFatSectors = mappedBuffer.getInt(64);
        this.firstDifatSector = mappedBuffer.getInt(68);
        // numDifatSectors at offset 72 - not needed for current implementation

        // Read root entry (index 0) to get mini stream location - zero-copy
        int rootEntryOffset = getSectorOffset(firstDirSector);
        int rootStartSector = mappedBuffer.getInt(rootEntryOffset + 116);
        long rootStreamSize = mappedBuffer.getLong(rootEntryOffset + 120);

        this.miniStreamStart = rootStreamSize > 0 ? rootStartSector : ENDOFCHAIN;
    }

    /**
     * Gets the file offset for a sector (zero-copy calculation).
     */
    private int getSectorOffset(int sectorIndex) {
        return HEADER_SIZE + sectorIndex * sectorSize;
    }

    /**
     * Gets a FAT sector location from DIFAT (zero-copy, no array).
     * DIFAT: first 109 entries in header at offset 76, rest in DIFAT sector chain.
     */
    private int getDifatEntry(int fatSectorNum) {
        if (fatSectorNum < 109) {
            // Read directly from header
            return mappedBuffer.getInt(76 + fatSectorNum * 4);
        }

        // Walk DIFAT sector chain (rare - only for files > ~6.7MB)
        int difatIndex = fatSectorNum - 109;
        int entriesPerDifatSector = fatEntriesPerSector - 1; // Last entry is next DIFAT sector
        int difatSectorNum = difatIndex / entriesPerDifatSector;
        int entryInDifatSector = difatIndex % entriesPerDifatSector;

        // Walk to the right DIFAT sector
        int difatSector = firstDifatSector;
        for (int i = 0; i < difatSectorNum && difatSector >= 0; i++) {
            int difatOffset = getSectorOffset(difatSector);
            difatSector = mappedBuffer.getInt(difatOffset + entriesPerDifatSector * 4);
        }

        if (difatSector < 0) {
            return ENDOFCHAIN;
        }

        return mappedBuffer.getInt(getSectorOffset(difatSector) + entryInDifatSector * 4);
    }

    /**
     * Reads a FAT entry on-demand from mapped buffer (zero-copy).
     */
    private int getFatEntry(int sectorIndex) {
        int fatSectorNum = sectorIndex / fatEntriesPerSector;
        int entryInSector = sectorIndex % fatEntriesPerSector;

        if (fatSectorNum >= numFatSectors) {
            return ENDOFCHAIN;
        }

        int fatSectorLocation = getDifatEntry(fatSectorNum);
        if (fatSectorLocation < 0) {
            return ENDOFCHAIN;
        }

        return mappedBuffer.getInt(getSectorOffset(fatSectorLocation) + entryInSector * 4);
    }

    /**
     * Reads a MiniFAT entry on-demand from mapped buffer (zero-copy).
     */
    private int getMiniFatEntry(int miniSectorIndex) {
        if (numMiniFatSectors == 0) {
            return ENDOFCHAIN;
        }

        int entriesPerSector = fatEntriesPerSector;
        int miniFatSectorNum = miniSectorIndex / entriesPerSector;
        int entryInSector = miniSectorIndex % entriesPerSector;

        // Walk MiniFAT sector chain to find the right sector
        int sector = firstMiniFatSector;
        for (int i = 0; i < miniFatSectorNum && sector != ENDOFCHAIN; i++) {
            sector = getFatEntry(sector);
        }

        if (sector < 0) {
            return ENDOFCHAIN;
        }

        int miniFatSectorOffset = getSectorOffset(sector);
        return mappedBuffer.getInt(miniFatSectorOffset + entryInSector * 4);
    }

    /**
     * Checks if directory entry at offset matches the given name (zero-copy comparison).
     */
    private boolean entryNameMatches(int offset, String name) {
        int nameLen = (mappedBuffer.getShort(offset + 64) & 0xFFFF) - 2;
        if (nameLen <= 0 || nameLen != name.length() * 2) {
            return false;
        }

        // Compare UTF-16LE bytes directly without allocation
        for (int i = 0; i < name.length(); i++) {
            char c = name.charAt(i);
            byte lo = mappedBuffer.get(offset + i * 2);
            byte hi = mappedBuffer.get(offset + i * 2 + 1);
            if ((c & 0xFF) != (lo & 0xFF) || ((c >> 8) & 0xFF) != (hi & 0xFF)) {
                return false;
            }
        }
        return true;
    }

    /**
     * Finds a directory entry by name (zero-copy search, lazy caching).
     */
    private @Nullable DirectoryEntry findDirectoryEntry(String name) {
        // Check cache first
        DirectoryEntry cached = switch (name) {
            case "EncryptedPackage" -> encryptedPackageEntry;
            case "EncryptionInfo" -> encryptionInfoEntry;
            default -> null;
        };
        if (cached != null) {
            return cached;
        }

        // Search directory sectors
        int dirSector = firstDirSector;
        while (dirSector >= 0) {
            int sectorOffset = getSectorOffset(dirSector);
            for (int i = 0; i < sectorSize / DIR_ENTRY_SIZE; i++) {
                int entryOffset = sectorOffset + i * DIR_ENTRY_SIZE;

                // Check if entry is valid
                byte type = mappedBuffer.get(entryOffset + 66);
                if (type == 0) {
                    continue;
                }

                // Check name match without allocation
                if (entryNameMatches(entryOffset, name)) {
                    DirectoryEntry entry = new DirectoryEntry(
                        name,
                        type,
                        mappedBuffer.getInt(entryOffset + 116),
                        mappedBuffer.getLong(entryOffset + 120)
                    );

                    // Cache common entries
                    if ("EncryptedPackage".equals(name)) {
                        encryptedPackageEntry = entry;
                    } else if ("EncryptionInfo".equals(name)) {
                        encryptionInfoEntry = entry;
                    }

                    return entry;
                }
            }
            dirSector = getFatEntry(dirSector);
        }
        return null;
    }

    /**
     * Returns a read-only ByteBuffer view of a stream (zero-copy for contiguous streams).
     *
     * @param name the stream name
     * @return the stream data, or null if not found
     */
    public @Nullable ByteBuffer readStreamAsBuffer(String name) {
        DirectoryEntry entry = findDirectoryEntry(name);
        return entry != null ? readStreamAsBuffer(entry) : null;
    }

    private byte[] readStream(DirectoryEntry entry) {
        if (entry.streamSize == 0) {
            return new byte[0];
        }

        // Root entry (type 5) stores the mini stream itself, always in regular sectors
        boolean useRegularSectors = entry.type == 5 || entry.streamSize >= miniStreamCutoffSize;

        // Check if stream is contiguous (can use zero-copy slice)
        if (useRegularSectors && isContiguousStream(entry)) {
            // Zero-copy: create a byte array backed by the mapped buffer data
            int offset = getSectorOffset(entry.startSector);
            byte[] result = new byte[(int) entry.streamSize];
            // Use bulk get from mapped buffer (optimized by JVM)
            mappedBuffer.duplicate().position(offset).get(result);
            return result;
        }

        // Non-contiguous stream: need to copy sector by sector
        byte[] result = new byte[(int) entry.streamSize];
        int destOffset = 0;

        if (!useRegularSectors) {
            // Read from mini stream (which is in regular sectors)
            int miniSector = entry.startSector;
            long remaining = entry.streamSize;
            while (miniSector >= 0 && remaining > 0) {
                // Find which regular sector contains this mini sector
                int miniStreamOffset = miniSector * miniSectorSize;
                int regularSectorIndex = miniStreamOffset / sectorSize;
                int offsetInSector = miniStreamOffset % sectorSize;

                // Walk mini stream's sector chain to find the right regular sector
                int regularSector = miniStreamStart;
                for (int i = 0; i < regularSectorIndex && regularSector >= 0; i++) {
                    regularSector = getFatEntry(regularSector);
                }

                if (regularSector >= 0) {
                    int srcOffset = getSectorOffset(regularSector) + offsetInSector;
                    int toRead = (int) Math.min(miniSectorSize, remaining);
                    for (int i = 0; i < toRead; i++) {
                        result[destOffset++] = mappedBuffer.get(srcOffset + i);
                    }
                    remaining -= toRead;
                }
                miniSector = getMiniFatEntry(miniSector);
            }
        } else {
            // Read from regular sectors
            int sector = entry.startSector;
            long remaining = entry.streamSize;
            while (sector >= 0 && remaining > 0) {
                int srcOffset = getSectorOffset(sector);
                int toRead = (int) Math.min(sectorSize, remaining);
                for (int i = 0; i < toRead; i++) {
                    result[destOffset++] = mappedBuffer.get(srcOffset + i);
                }
                remaining -= toRead;
                sector = getFatEntry(sector);
            }
        }

        return result;
    }

    /**
     * Returns a zero-copy ByteBuffer slice for contiguous streams.
     */
    private ByteBuffer readStreamAsBuffer(DirectoryEntry entry) {
        if (entry.streamSize == 0) {
            return ByteBuffer.allocate(0);
        }

        boolean useRegularSectors = entry.type == 5 || entry.streamSize >= miniStreamCutoffSize;

        if (useRegularSectors && isContiguousStream(entry)) {
            // Zero-copy: return a slice of the mapped buffer
            int offset = getSectorOffset(entry.startSector);
            ByteBuffer slice = mappedBuffer.duplicate();
            slice.position(offset);
            slice.limit(offset + (int) entry.streamSize);
            return slice.slice().order(ByteOrder.LITTLE_ENDIAN).asReadOnlyBuffer();
        }

        // Non-contiguous: fall back to copy
        byte[] data = readStream(entry);
        return ByteBuffer.wrap(data).order(ByteOrder.LITTLE_ENDIAN).asReadOnlyBuffer();
    }

    /**
     * Checks if a stream's sectors are contiguous (enables zero-copy).
     */
    private boolean isContiguousStream(DirectoryEntry entry) {
        if (entry.streamSize <= sectorSize) {
            return true; // Single sector is always contiguous
        }

        int numSectors = (int) ((entry.streamSize + sectorSize - 1) / sectorSize);
        int sector = entry.startSector;

        for (int i = 0; i < numSectors - 1; i++) {
            int nextSector = getFatEntry(sector);
            if (nextSector != sector + 1) {
                return false; // Not contiguous
            }
            sector = nextSector;
        }
        return true;
    }

    /**
     * Checks if this file is an encrypted Office document.
     */
    public boolean isEncrypted() {
        return findDirectoryEntry("EncryptionInfo") != null
            || findDirectoryEntry("EncryptedPackage") != null;
    }

    /**
     * Gets the encrypted package as a zero-copy ByteBuffer (if contiguous), or null if not found.
     * This avoids allocating memory for large encrypted files.
     */
    public @Nullable ByteBuffer getEncryptedPackageBuffer() {
        return readStreamAsBuffer("EncryptedPackage");
    }

    /**
     * Gets the encryption info as a zero-copy ByteBuffer (if contiguous), or null if not found.
     */
    public @Nullable ByteBuffer getEncryptionInfoBuffer() {
        return readStreamAsBuffer("EncryptionInfo");
    }

    @Override
    public void close() throws IOException {
        channel.close();
    }

    private record DirectoryEntry(String name, byte type, int startSector, long streamSize) {}
}
