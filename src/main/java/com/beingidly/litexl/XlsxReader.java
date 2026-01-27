package com.beingidly.litexl;

import com.beingidly.litexl.crypto.AgileDecryptor;
import com.beingidly.litexl.crypto.CfbReader;

import org.jspecify.annotations.Nullable;

import java.io.*;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.security.GeneralSecurityException;
import java.util.*;

/**
 * Reads XLSX files.
 *
 * <p>This class implements {@link Closeable} and should be used with try-with-resources:</p>
 * <pre>{@code
 * try (XlsxReader reader = new XlsxReader(path)) {
 *     Workbook wb = reader.read();
 * }
 * }</pre>
 */
final class XlsxReader implements Closeable {

    private final Path path;
    private final List<String> sharedStrings = new ArrayList<>();
    private @Nullable ZipReader zip;

    /**
     * Creates a new XLSX reader for the given file.
     *
     * @param path the path to the XLSX file
     */
    public XlsxReader(Path path) {
        this.path = path;
    }

    /**
     * Reads the workbook without a password.
     *
     * @return the parsed workbook
     * @throws IOException if an I/O error occurs
     */
    public Workbook read() throws IOException {
        return read(null);
    }

    /**
     * Reads the workbook with an optional password.
     *
     * @param password the password for encrypted files, or null
     * @return the parsed workbook
     * @throws IOException if an I/O error occurs
     */
    public Workbook read(@Nullable String password) throws IOException {
        // Check if this is an encrypted CFB file
        if (isEncryptedFile()) {
            return readEncrypted(password);
        }

        this.zip = new ZipReader(path);
        Workbook wb = Workbook.create();

        // Read shared strings first
        readSharedStrings();

        // Read styles
        readStyles();

        // Read workbook.xml to get sheet names and order
        List<SheetInfo> sheetInfos = readWorkbook();

        // Read each sheet
        for (SheetInfo info : sheetInfos) {
            Sheet sheet = new Sheet(info.name(), info.index());
            readSheet(info.path(), sheet);
            wb.addSheet(sheet);
        }

        // Populate shared strings in workbook
        for (String s : sharedStrings) {
            wb.addSharedString(s);
        }

        return wb;
    }

    private void readSharedStrings() throws IOException {
        assert zip != null : "zip must be initialized";
        InputStream is = zip.getEntry("xl/sharedStrings.xml");
        if (is == null) {
            return;
        }

        try (is) {
            XmlReader xml = new XmlReader(is);
            while (xml.hasNext()) {
                XmlReader.Event event = xml.next();
                if (event == XmlReader.Event.START_ELEMENT && "t".equals(xml.getLocalName())) {
                    sharedStrings.add(xml.getElementText());
                }
            }
        }
    }

    private void readStyles() throws IOException {
        assert zip != null : "zip must be initialized";
        InputStream is = zip.getEntry("xl/styles.xml");
        if (is == null) {
            return;
        }

        // TODO: Parse styles properly - for now we just skip
        is.close();
    }

    private List<SheetInfo> readWorkbook() throws IOException {
        assert zip != null : "zip must be initialized";
        List<SheetInfo> sheets = new ArrayList<>();
        Map<String, String> relsMap = readWorkbookRelationships();

        InputStream is = zip.getEntry("xl/workbook.xml");
        if (is == null) {
            throw new CorruptFileException("Missing xl/workbook.xml");
        }

        try (is) {
            XmlReader xml = new XmlReader(is);
            int index = 0;
            while (xml.hasNext()) {
                XmlReader.Event event = xml.next();
                if (event == XmlReader.Event.START_ELEMENT && "sheet".equals(xml.getLocalName())) {
                    String name = xml.getAttributeValue("name");
                    if (name == null) {
                        continue; // Skip sheets without name (corrupt file)
                    }
                    String rId = xml.getAttributeValue("r:id");
                    if (rId == null) {
                        rId = xml.getAttributeValue("id");
                    }
                    if (rId != null) {
                        String target = relsMap.get(rId);
                        if (target != null) {
                            sheets.add(new SheetInfo(name, index++, "xl/" + target));
                        }
                    }
                }
            }
        }

        return sheets;
    }

    private Map<String, String> readWorkbookRelationships() throws IOException {
        assert zip != null : "zip must be initialized";
        Map<String, String> rels = new HashMap<>();

        InputStream is = zip.getEntry("xl/_rels/workbook.xml.rels");
        if (is == null) {
            return rels;
        }

        try (is) {
            XmlReader xml = new XmlReader(is);
            while (xml.hasNext()) {
                XmlReader.Event event = xml.next();
                if (event == XmlReader.Event.START_ELEMENT && "Relationship".equals(xml.getLocalName())) {
                    String id = xml.getAttributeValue("Id");
                    String target = xml.getAttributeValue("Target");
                    if (id != null && target != null) {
                        rels.put(id, target);
                    }
                }
            }
        }

        return rels;
    }

    private void readSheet(String sheetPath, Sheet sheet) throws IOException {
        assert zip != null : "zip must be initialized";
        InputStream is = zip.getEntry(sheetPath);
        if (is == null) {
            throw new CorruptFileException("Missing sheet: " + sheetPath);
        }

        try (is) {
            XmlReader xml = new XmlReader(is);
            String currentRef = null;
            String currentType = null;
            int currentStyle = 0;
            boolean inInlineStr = false;

            while (xml.hasNext()) {
                XmlReader.Event event = xml.next();

                if (event == XmlReader.Event.START_ELEMENT) {
                    String name = xml.getLocalName();

                    if ("c".equals(name)) {
                        // Cell element
                        currentRef = xml.getAttributeValue("r");
                        currentType = xml.getAttributeValue("t");
                        String styleStr = xml.getAttributeValue("s");
                        currentStyle = styleStr != null ? Integer.parseInt(styleStr) : 0;
                    } else if ("is".equals(name)) {
                        // Inline string container
                        inInlineStr = true;
                    } else if ("t".equals(name) && inInlineStr) {
                        // Inline string text
                        assert currentRef != null : "currentRef must be set inside cell";
                        String value = xml.getElementText();
                        int[] coords = CellRefUtil.parseRef(currentRef);
                        Cell cell = sheet.cell(coords[0], coords[1]);
                        cell.set(value);
                        cell.style(currentStyle);
                    } else if ("v".equals(name)) {
                        // Value element
                        assert currentRef != null : "currentRef must be set inside cell";
                        String value = xml.getElementText();
                        int[] coords = CellRefUtil.parseRef(currentRef);
                        Cell cell = sheet.cell(coords[0], coords[1]);
                        cell.style(currentStyle);

                        if (currentType == null) {
                            // Number or date (default)
                            double num = Double.parseDouble(value);
                            cell.set(num);
                        } else {
                            switch (currentType) {
                                case "s" -> {
                                    // Shared string
                                    int idx = Integer.parseInt(value);
                                    if (idx < sharedStrings.size()) {
                                        cell.set(sharedStrings.get(idx));
                                    }
                                }
                                case "b" -> cell.set("1".equals(value)); // Boolean
                                case "e" -> cell.setValue(new CellValue.Error(value)); // Error
                                default -> {
                                    // Other types - treat as number
                                    double num = Double.parseDouble(value);
                                    cell.set(num);
                                }
                            }
                        }
                    } else if ("f".equals(name)) {
                        // Formula element
                        assert currentRef != null : "currentRef must be set inside cell";
                        String formula = xml.getElementText();
                        int[] coords = CellRefUtil.parseRef(currentRef);
                        Cell cell = sheet.cell(coords[0], coords[1]);
                        cell.setFormula(formula);
                        cell.style(currentStyle);
                    } else if ("mergeCell".equals(name)) {
                        // Merged cell
                        String ref = xml.getAttributeValue("ref");
                        if (ref != null) {
                            CellRange range = CellRange.parse(ref);
                            sheet.merge(range);
                        }
                    }
                } else if (event == XmlReader.Event.END_ELEMENT) {
                    String name = xml.getLocalName();
                    if ("c".equals(name)) {
                        currentRef = null;
                        currentType = null;
                        currentStyle = 0;
                        inInlineStr = false;
                    } else if ("is".equals(name)) {
                        inInlineStr = false;
                    }
                }
            }
        }
    }

    private record SheetInfo(String name, int index, String path) {}

    /**
     * Checks if the file is an encrypted CFB (OLE2) file.
     */
    private boolean isEncryptedFile() throws IOException {
        byte[] header = new byte[8];
        try (InputStream is = Files.newInputStream(path)) {
            if (is.read(header) != 8) {
                return false;
            }
        }

        // CFB signature: D0 CF 11 E0 A1 B1 1A E1
        return header[0] == (byte) 0xD0 &&
               header[1] == (byte) 0xCF &&
               header[2] == (byte) 0x11 &&
               header[3] == (byte) 0xE0 &&
               header[4] == (byte) 0xA1 &&
               header[5] == (byte) 0xB1 &&
               header[6] == (byte) 0x1A &&
               header[7] == (byte) 0xE1;
    }

    /**
     * Reads an encrypted workbook.
     */
    private Workbook readEncrypted(@Nullable String password) throws IOException {
        if (password == null) {
            throw new InvalidPasswordException("Password required for encrypted file");
        }

        try (CfbReader cfb = new CfbReader(path)) {
            if (!cfb.isEncrypted()) {
                throw new CorruptFileException("CFB file is not an encrypted Office document");
            }

            ByteBuffer encryptionInfoData = cfb.getEncryptionInfoBuffer();
            ByteBuffer encryptedPackage = cfb.getEncryptedPackageBuffer();

            if (encryptionInfoData == null || encryptedPackage == null) {
                throw new CorruptFileException("Missing encryption streams");
            }

            // Parse encryption info and decrypt
            AgileDecryptor decryptor = new AgileDecryptor(password);
            decryptor.parseEncryptionInfo(encryptionInfoData, encryptedPackage);
            ByteBuffer decryptedData;
            try {
                decryptedData = decryptor.decrypt();
            } catch (GeneralSecurityException e) {
                throw new InvalidPasswordException("Decryption failed", e);
            }

            // Now read the decrypted XLSX data from memory
            return readFromByteBuffer(decryptedData);
        }
    }

    /**
     * Reads a workbook from a ByteBuffer (decrypted XLSX data).
     */
    private Workbook readFromByteBuffer(ByteBuffer data) throws IOException {
        Path tempFile = Files.createTempFile("litexl-", ".xlsx");
        try {
            try (FileChannel channel = FileChannel.open(tempFile, StandardOpenOption.WRITE)) {
                while (data.hasRemaining()) {
                    int written = channel.write(data);
                    if (written == 0) {
                        throw new IOException("Failed to write data to temp file");
                    }
                }
            }
            try (XlsxReader tempReader = new XlsxReader(tempFile)) {
                return tempReader.read();
            }
        } finally {
            Files.deleteIfExists(tempFile);
        }
    }

    @Override
    public void close() throws IOException {
        if (zip != null) {
            zip.close();
            zip = null;
        }
    }
}
