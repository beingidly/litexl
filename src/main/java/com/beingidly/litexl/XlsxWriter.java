package com.beingidly.litexl;

import com.beingidly.litexl.crypto.EncryptionOptions;
import com.beingidly.litexl.crypto.SheetHasher;
import com.beingidly.litexl.format.*;
import com.beingidly.litexl.style.*;
import org.jspecify.annotations.Nullable;

import java.io.*;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.security.GeneralSecurityException;
import java.util.Map;
import java.util.Objects;

/**
 * Writes XLSX files.
 *
 * <p>This class implements {@link Closeable} and should be used with try-with-resources:</p>
 * <pre>{@code
 * try (XlsxWriter writer = new XlsxWriter(workbook, path)) {
 *     writer.write();
 * }
 * }</pre>
 *
 * <p>For encrypted files:</p>
 * <pre>{@code
 * try (XlsxWriter writer = new XlsxWriter(workbook, path)
 *         .withEncryption(EncryptionOptions.aes256("password"))) {
 *     writer.write();
 * }
 * }</pre>
 */
final class XlsxWriter implements Closeable {

    private static final String NS_SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static final String NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static final String NS_PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";

    private final Workbook workbook;
    private final @Nullable Path path;
    private final @Nullable OutputStream outputStream;
    private @Nullable EncryptionOptions encryptionOptions;

    /**
     * Creates a new XLSX writer for the given workbook and output path.
     *
     * @param workbook the workbook to write
     * @param path the output file path
     */
    public XlsxWriter(Workbook workbook, Path path) {
        this.workbook = workbook;
        this.path = path;
        this.outputStream = null;
    }

    /**
     * Creates a new XLSX writer for the given workbook and output stream.
     *
     * @param workbook the workbook to write
     * @param outputStream the output stream to write to
     */
    public XlsxWriter(Workbook workbook, OutputStream outputStream) {
        this.workbook = workbook;
        this.path = null;
        this.outputStream = outputStream;
    }

    /**
     * Enables encryption for the output file.
     *
     * @param options the encryption options
     * @return this writer for method chaining
     */
    public XlsxWriter withEncryption(EncryptionOptions options) {
        this.encryptionOptions = options;
        return this;
    }

    /**
     * Writes the workbook to the output file or stream.
     *
     * @throws IOException if an I/O error occurs
     */
    public void write() throws IOException {
        if (encryptionOptions != null) {
            writeEncrypted();
            return;
        }

        try (ZipWriter zip = createZipWriter()) {
            writeContentTypes(zip);
            writeRootRels(zip);
            writeWorkbookRels(zip);
            writeWorkbook(zip);
            writeStyles(zip);

            for (int i = 0; i < workbook.sheetCount(); i++) {
                writeSheet(zip, Objects.requireNonNull(workbook.getSheet(i)), i + 1);
            }
        }
    }

    private ZipWriter createZipWriter() throws IOException {
        if (outputStream != null) {
            return new ZipWriter(outputStream);
        }
        assert path != null : "path must be set when outputStream is null";
        return new ZipWriter(path);
    }

    private void writeContentTypes(ZipWriter zip) throws IOException {
        try (OutputStream os = zip.newEntry("[Content_Types].xml");
             XmlWriter xml = new XmlWriter(os)) {

            xml.startDocument();
            xml.startElement("Types");
            xml.attribute("xmlns", NS_CONTENT_TYPES);

            // Defaults
            xml.emptyElement("Default");
            xml.attribute("Extension", "rels");
            xml.attribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");

            xml.emptyElement("Default");
            xml.attribute("Extension", "xml");
            xml.attribute("ContentType", "application/xml");

            // Overrides
            xml.emptyElement("Override");
            xml.attribute("PartName", "/xl/workbook.xml");
            xml.attribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");

            xml.emptyElement("Override");
            xml.attribute("PartName", "/xl/styles.xml");
            xml.attribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");

            for (int i = 0; i < workbook.sheetCount(); i++) {
                xml.emptyElement("Override");
                xml.attribute("PartName", "/xl/worksheets/sheet" + (i + 1) + ".xml");
                xml.attribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            }

            xml.endElement(); // Types
            xml.endDocument();
        }
    }

    private void writeRootRels(ZipWriter zip) throws IOException {
        try (OutputStream os = zip.newEntry("_rels/.rels");
             XmlWriter xml = new XmlWriter(os)) {

            xml.startDocument();
            xml.startElement("Relationships");
            xml.attribute("xmlns", NS_PACKAGE_RELS);

            xml.emptyElement("Relationship");
            xml.attribute("Id", "rId1");
            xml.attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            xml.attribute("Target", "xl/workbook.xml");

            xml.endElement();
            xml.endDocument();
        }
    }

    private void writeWorkbookRels(ZipWriter zip) throws IOException {
        try (OutputStream os = zip.newEntry("xl/_rels/workbook.xml.rels");
             XmlWriter xml = new XmlWriter(os)) {

            xml.startDocument();
            xml.startElement("Relationships");
            xml.attribute("xmlns", NS_PACKAGE_RELS);

            int rId = 1;

            // Sheets
            for (int i = 0; i < workbook.sheetCount(); i++) {
                xml.emptyElement("Relationship");
                xml.attribute("Id", "rId" + rId++);
                xml.attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                xml.attribute("Target", "worksheets/sheet" + (i + 1) + ".xml");
            }

            // Styles
            xml.emptyElement("Relationship");
            xml.attribute("Id", "rId" + rId);
            xml.attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            xml.attribute("Target", "styles.xml");

            xml.endElement();
            xml.endDocument();
        }
    }

    private void writeWorkbook(ZipWriter zip) throws IOException {
        try (OutputStream os = zip.newEntry("xl/workbook.xml");
             XmlWriter xml = new XmlWriter(os)) {

            xml.startDocument();
            xml.startElement("workbook");
            xml.attribute("xmlns", NS_SPREADSHEETML);
            xml.attribute("xmlns:r", NS_RELATIONSHIPS);

            xml.startElement("sheets");
            for (int i = 0; i < workbook.sheetCount(); i++) {
                Sheet sheet = workbook.getSheet(i);
                assert sheet != null;

                xml.emptyElement("sheet");
                xml.attribute("name", sheet.name());
                xml.attribute("sheetId", String.valueOf(i + 1));
                xml.attribute("r:id", "rId" + (i + 1));
            }
            xml.endElement(); // sheets

            xml.endElement(); // workbook
            xml.endDocument();
        }
    }

    private void writeStyles(ZipWriter zip) throws IOException {
        try (OutputStream os = zip.newEntry("xl/styles.xml")) {
            StylesXml.write(os, workbook.styles());
        }
    }

    private void writeSheet(ZipWriter zip, Sheet sheet, int sheetNum) throws IOException {
        try (OutputStream os = zip.newEntry("xl/worksheets/sheet" + sheetNum + ".xml");
             XmlWriter xml = new XmlWriter(os)) {

            xml.startDocument();
            xml.startElement("worksheet");
            xml.attribute("xmlns", NS_SPREADSHEETML);
            xml.attribute("xmlns:r", NS_RELATIONSHIPS);

            // Column widths (must come before sheetData per OOXML spec)
            Map<Integer, Double> colWidths = sheet.columnWidths();
            if (!colWidths.isEmpty()) {
                xml.startElement("cols");
                for (Map.Entry<Integer, Double> entry : colWidths.entrySet()) {
                    int col = entry.getKey();
                    double width = entry.getValue();
                    xml.emptyElement("col");
                    xml.attribute("min", String.valueOf(col + 1));
                    xml.attribute("max", String.valueOf(col + 1));
                    xml.attribute("width", String.valueOf(width));
                    xml.attribute("customWidth", "1");
                }
                xml.endElement(); // cols
            }

            // Sheet data
            xml.startElement("sheetData");

            try {
                sheet.forEachRow(row -> {
                    try {
                        xml.startElement("row");
                        xml.attribute("r", String.valueOf(row.rowNum() + 1));

                        if (row.hasCustomHeight()) {
                            xml.attribute("ht", String.valueOf(row.height()));
                            xml.attribute("customHeight", "1");
                        }

                        for (Map.Entry<Integer, Cell> cellEntry : row.cells().entrySet()) {
                            Cell cell = cellEntry.getValue();
                            writeCell(xml, cell, row.rowNum());
                        }

                        xml.endElement(); // row
                    } catch (IOException e) {
                        throw new UncheckedIOException(e);
                    }
                    return true;
                });
            } catch (UncheckedIOException e) {
                throw e.getCause();
            }

            xml.endElement(); // sheetData

            // Sheet protection
            if (sheet.isProtected()) {
                writeSheetProtection(xml, sheet);
            }

            // AutoFilter (must come before mergeCells per OOXML spec)
            if (sheet.autoFilter() != null) {
                writeAutoFilter(xml, sheet.autoFilter());
            }

            // Merged cells
            if (!sheet.mergedCells().isEmpty()) {
                xml.startElement("mergeCells");
                xml.attribute("count", String.valueOf(sheet.mergedCells().size()));
                for (Sheet.MergedRegion merge : sheet.mergedCells()) {
                    xml.emptyElement("mergeCell");
                    xml.attribute("ref", merge.toRange().toRef());
                }
                xml.endElement();
            }

            // Conditional formatting
            if (!sheet.conditionalFormats().isEmpty()) {
                for (ConditionalFormat cf : sheet.conditionalFormats()) {
                    writeConditionalFormat(xml, cf);
                }
            }

            // Data validations
            if (!sheet.validations().isEmpty()) {
                xml.startElement("dataValidations");
                xml.attribute("count", String.valueOf(sheet.validations().size()));
                for (DataValidation dv : sheet.validations()) {
                    writeDataValidation(xml, dv);
                }
                xml.endElement();
            }

            xml.endElement(); // worksheet
            xml.endDocument();
        }
    }

    private void writeCell(XmlWriter xml, Cell cell, int row) throws IOException {
        if (cell.type() == CellType.EMPTY) {
            return;
        }

        String ref = CellRefUtil.toRef(row, cell.column());

        xml.startElement("c");
        xml.attribute("r", ref);

        if (cell.styleId() > 0) {
            xml.attribute("s", String.valueOf(cell.styleId()));
        }

        switch (cell.value()) {
            case CellValue.Text t -> {
                xml.attribute("t", "inlineStr");
                xml.startElement("is");
                xml.startElement("t");
                String text = t.value();
                if (!text.isEmpty() && (Character.isWhitespace(text.charAt(0))
                        || Character.isWhitespace(text.charAt(text.length() - 1)))) {
                    xml.attribute("xml:space", "preserve");
                }
                xml.text(text);
                xml.endElement(); // t
                xml.endElement(); // is
            }
            case CellValue.Number n -> {
                xml.startElement("v");
                xml.text(String.valueOf(n.value()));
                xml.endElement();
            }
            case CellValue.Bool b -> {
                xml.attribute("t", "b");
                xml.startElement("v");
                xml.text(b.value() ? "1" : "0");
                xml.endElement();
            }
            case CellValue.Date d -> {
                xml.startElement("v");
                xml.text(String.valueOf(ExcelDateUtil.toExcelDate(d.value())));
                xml.endElement();
            }
            case CellValue.Formula f -> {
                xml.startElement("f");
                xml.text(f.expression());
                xml.endElement();
            }
            case CellValue.Error e -> {
                xml.attribute("t", "e");
                xml.startElement("v");
                xml.text(e.code());
                xml.endElement();
            }
            case CellValue.Empty _ -> {
                // Skip
            }
        }

        xml.endElement(); // c
    }

    private void writeConditionalFormat(XmlWriter xml, ConditionalFormat cf) throws IOException {
        xml.startElement("conditionalFormatting");
        xml.attribute("sqref", cf.range().toRef());

        xml.startElement("cfRule");
        xml.attribute("type", cfTypeName(cf.type()));
        xml.attribute("priority", "1");

        if (cf.operator() != ConditionalFormat.Operator.NONE) {
            xml.attribute("operator", cfOperatorName(cf.operator()));
        }

        if (cf.styleId() > 0) {
            xml.attribute("dxfId", String.valueOf(cf.styleId() - 1));
        }

        if (cf.formula1() != null) {
            xml.startElement("formula");
            xml.text(cf.formula1());
            xml.endElement();
        }

        if (cf.formula2() != null) {
            xml.startElement("formula");
            xml.text(cf.formula2());
            xml.endElement();
        }

        xml.endElement(); // cfRule
        xml.endElement(); // conditionalFormatting
    }

    private String cfTypeName(ConditionalFormat.Type type) {
        return switch (type) {
            case CELL_VALUE -> "cellIs";
            case EXPRESSION -> "expression";
            case COLOR_SCALE -> "colorScale";
            case DATA_BAR -> "dataBar";
            case ICON_SET -> "iconSet";
            case TOP_BOTTOM -> "top10";
            case ABOVE_AVERAGE -> "aboveAverage";
            case DUPLICATE_VALUES -> "duplicateValues";
            case UNIQUE_VALUES -> "uniqueValues";
            case CONTAINS_TEXT -> "containsText";
            case NOT_CONTAINS_TEXT -> "notContainsText";
            case BEGINS_WITH -> "beginsWith";
            case ENDS_WITH -> "endsWith";
            case CONTAINS_BLANKS -> "containsBlanks";
            case CONTAINS_ERRORS -> "containsErrors";
        };
    }

    private String cfOperatorName(ConditionalFormat.Operator op) {
        return switch (op) {
            case NONE -> "";
            case LESS_THAN -> "lessThan";
            case LESS_THAN_OR_EQUAL -> "lessThanOrEqual";
            case EQUAL -> "equal";
            case NOT_EQUAL -> "notEqual";
            case GREATER_THAN_OR_EQUAL -> "greaterThanOrEqual";
            case GREATER_THAN -> "greaterThan";
            case BETWEEN -> "between";
            case NOT_BETWEEN -> "notBetween";
        };
    }

    private void writeDataValidation(XmlWriter xml, DataValidation dv) throws IOException {
        xml.startElement("dataValidation");
        xml.attribute("sqref", dv.range().toRef());
        xml.attribute("type", dvTypeName(dv.type()));

        DataValidation.Operator dvOp = dv.operator();
        if (dvOp != null && dvOp != DataValidation.Operator.BETWEEN) {
            xml.attribute("operator", dvOperatorName(dvOp));
        }

        if (dv.showDropdown() && dv.type() == DataValidation.Type.LIST) {
            xml.attribute("showDropDown", "0"); // 0 = show, 1 = hide (counterintuitive!)
        }

        xml.attribute("allowBlank", "1");
        xml.attribute("showInputMessage", "1");
        xml.attribute("showErrorMessage", "1");

        if (dv.errorTitle() != null) {
            xml.attribute("errorTitle", dv.errorTitle());
        }
        if (dv.errorMessage() != null) {
            xml.attribute("error", dv.errorMessage());
        }

        if (dv.formula1() != null) {
            xml.startElement("formula1");
            xml.text(dv.formula1());
            xml.endElement();
        }

        if (dv.formula2() != null) {
            xml.startElement("formula2");
            xml.text(dv.formula2());
            xml.endElement();
        }

        xml.endElement();
    }

    private String dvTypeName(DataValidation.Type type) {
        return switch (type) {
            case ANY -> "none";
            case WHOLE -> "whole";
            case DECIMAL -> "decimal";
            case LIST -> "list";
            case DATE -> "date";
            case TIME -> "time";
            case TEXT_LENGTH -> "textLength";
            case CUSTOM -> "custom";
        };
    }

    private String dvOperatorName(DataValidation.Operator op) {
        return switch (op) {
            case BETWEEN -> "between";
            case NOT_BETWEEN -> "notBetween";
            case EQUAL -> "equal";
            case NOT_EQUAL -> "notEqual";
            case GREATER_THAN -> "greaterThan";
            case LESS_THAN -> "lessThan";
            case GREATER_THAN_OR_EQUAL -> "greaterThanOrEqual";
            case LESS_THAN_OR_EQUAL -> "lessThanOrEqual";
        };
    }

    private void writeAutoFilter(XmlWriter xml, AutoFilter af) throws IOException {
        xml.startElement("autoFilter");
        xml.attribute("ref", af.range().toRef());

        for (AutoFilter.FilterColumn fc : af.columns()) {
            xml.startElement("filterColumn");
            xml.attribute("colId", String.valueOf(fc.columnIndex()));

            if (fc.custom() != null) {
                AutoFilter.CustomFilter cf = fc.custom();
                xml.startElement("customFilters");
                if (!cf.and()) {
                    xml.attribute("and", "0");
                }

                xml.emptyElement("customFilter");
                xml.attribute("operator", filterOperatorName(cf.op1()));
                xml.attribute("val", cf.val1());

                if (cf.op2() != null && cf.val2() != null) {
                    xml.emptyElement("customFilter");
                    xml.attribute("operator", filterOperatorName(cf.op2()));
                    xml.attribute("val", cf.val2());
                }

                xml.endElement(); // customFilters
            } else if (!fc.values().isEmpty()) {
                xml.startElement("filters");
                for (String val : fc.values()) {
                    xml.emptyElement("filter");
                    xml.attribute("val", val);
                }
                xml.endElement();
            }

            xml.endElement(); // filterColumn
        }

        xml.endElement(); // autoFilter
    }

    private String filterOperatorName(AutoFilter.CustomFilter.Operator op) {
        return switch (op) {
            case EQUAL -> "equal";
            case NOT_EQUAL -> "notEqual";
            case GREATER_THAN -> "greaterThan";
            case GREATER_THAN_OR_EQUAL -> "greaterThanOrEqual";
            case LESS_THAN -> "lessThan";
            case LESS_THAN_OR_EQUAL -> "lessThanOrEqual";
        };
    }

    private void writeSheetProtection(XmlWriter xml, Sheet sheet) throws IOException {
        com.beingidly.litexl.crypto.SheetProtection prot = sheet.protection();
        if (prot == null) {
            return;
        }

        xml.startElement("sheetProtection");
        xml.attribute("sheet", "1");

        // Use pre-computed hash info from protection manager
        SheetHasher.SheetProtectionInfo hashInfo = sheet.protectionManager().passwordInfo();
        if (hashInfo != null) {
            xml.attribute("algorithmName", hashInfo.algorithmName());
            xml.attribute("hashValue", hashInfo.hashValue());
            xml.attribute("saltValue", hashInfo.saltValue());
            xml.attribute("spinCount", String.valueOf(hashInfo.spinCount()));
        }

        // Protection options (Excel uses inverse logic - 0 means allowed)
        xml.attribute("objects", "1");
        xml.attribute("scenarios", "1");

        if (!prot.selectLockedCells()) {
            xml.attribute("selectLockedCells", "1");
        }
        if (!prot.selectUnlockedCells()) {
            xml.attribute("selectUnlockedCells", "1");
        }
        if (!prot.formatCells()) {
            xml.attribute("formatCells", "1");
        }
        if (!prot.formatColumns()) {
            xml.attribute("formatColumns", "1");
        }
        if (!prot.formatRows()) {
            xml.attribute("formatRows", "1");
        }
        if (!prot.insertRows()) {
            xml.attribute("insertRows", "1");
        }
        if (!prot.insertColumns()) {
            xml.attribute("insertColumns", "1");
        }
        if (!prot.deleteRows()) {
            xml.attribute("deleteRows", "1");
        }
        if (!prot.deleteColumns()) {
            xml.attribute("deleteColumns", "1");
        }
        if (!prot.sort()) {
            xml.attribute("sort", "1");
        }
        if (!prot.autoFilter()) {
            xml.attribute("autoFilter", "1");
        }
        if (!prot.pivotTables()) {
            xml.attribute("pivotTables", "1");
        }

        xml.endElement();
    }

    /**
     * Writes an encrypted workbook using MappedByteBuffer.
     * No heap/direct buffer allocation - uses OS page cache directly.
     */
    private void writeEncrypted() throws IOException {
        // First, write the unencrypted XLSX to a temp file
        Path xlsxTempFile = Files.createTempFile("litexl-", ".xlsx");
        try {
            try (ZipWriter zip = new ZipWriter(xlsxTempFile)) {
                writeContentTypes(zip);
                writeRootRels(zip);
                writeWorkbookRels(zip);
                writeWorkbook(zip);
                writeStyles(zip);
                for (int i = 0; i < workbook.sheetCount(); i++) {
                    writeSheet(zip, Objects.requireNonNull(workbook.getSheet(i)), i + 1);
                }
            }

            long xlsxSize = Files.size(xlsxTempFile);

            try {
                if (path != null) {
                    // Path-based: write directly to final file via FileChannel
                    writeAsCfbDirect(xlsxTempFile, xlsxSize, path);
                } else {
                    // OutputStream-based: write to temp CFB file, then transferTo
                    assert outputStream != null;
                    writeAsCfbToStream(xlsxTempFile, xlsxSize, outputStream);
                }
            } catch (GeneralSecurityException e) {
                throw new IOException("Encryption failed", e);
            }

        } finally {
            Files.deleteIfExists(xlsxTempFile);
        }
    }

    /**
     * Writes encrypted CFB directly to the destination file via FileChannel + MappedByteBuffer.
     */
    private void writeAsCfbDirect(Path xlsxTempFile, long xlsxSize, Path destPath) throws IOException, GeneralSecurityException {
        EncryptionData encryptionData = generateEncryptionData();

        try (InputStream xlsxIn = Files.newInputStream(xlsxTempFile);
             FileChannel fc = FileChannel.open(destPath, StandardOpenOption.READ, StandardOpenOption.WRITE, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING);
             CfbWriter cfbWriter = new CfbWriter(
                 fc,
                 encryptionData.encryptionInfo(),
                 encryptionData.encryptionKey(),
                 encryptionData.keyDataSalt(),
                 encryptionData.hmacKey()
             )) {

            cfbWriter.writeEncrypted(xlsxIn, xlsxSize);
        }
    }

    /**
     * Writes encrypted CFB to a temp file, then transfers to OutputStream.
     */
    private void writeAsCfbToStream(Path xlsxTempFile, long xlsxSize, OutputStream out) throws IOException, GeneralSecurityException {
        EncryptionData encryptionData = generateEncryptionData();

        Path cfbTempFile = Files.createTempFile("litexl-cfb-", ".tmp");
        try {
            // Write to temp CFB file via MappedByteBuffer
            try (InputStream xlsxIn = Files.newInputStream(xlsxTempFile);
                 FileChannel fc = FileChannel.open(cfbTempFile, StandardOpenOption.WRITE, StandardOpenOption.READ, StandardOpenOption.CREATE);
                 CfbWriter cfbWriter = new CfbWriter(
                     fc,
                     encryptionData.encryptionInfo(),
                     encryptionData.encryptionKey(),
                     encryptionData.keyDataSalt(),
                     encryptionData.hmacKey()
                 )) {

                cfbWriter.writeEncrypted(xlsxIn, xlsxSize);
            }

            // Transfer to OutputStream via NIO (zero-copy when possible)
            try (FileChannel fc = FileChannel.open(cfbTempFile, StandardOpenOption.READ)) {
                fc.transferTo(0, fc.size(), Channels.newChannel(out));
            }
        } finally {
            Files.deleteIfExists(cfbTempFile);
        }
    }

    /**
     * Generates encryption components for CFB/Agile encryption.
     * The final HMAC value is computed after EncryptedPackage bytes are written.
     */
    private EncryptionData generateEncryptionData() throws GeneralSecurityException, IOException {
        assert encryptionOptions != null : "encryptionOptions must be set";

        java.security.SecureRandom random = new java.security.SecureRandom();
        int keyBits = encryptionOptions.algorithm() == EncryptionOptions.Algorithm.AES_256 ? 256 : 128;

        // Two different salts per MS-OFFCRYPTO spec
        byte[] keyDataSalt = new byte[16];
        byte[] encryptedKeySalt = new byte[16];
        random.nextBytes(keyDataSalt);
        random.nextBytes(encryptedKeySalt);

        byte[] encryptionKey = new byte[keyBits / 8];
        random.nextBytes(encryptionKey);

        // Derive keys from password
        byte[] keyDerivedKey = com.beingidly.litexl.crypto.KeyDerivation.deriveKey(
            encryptionOptions.password(), encryptedKeySalt, encryptionOptions.spinCount(), keyBits,
            com.beingidly.litexl.crypto.KeyDerivation.BLOCK_KEY_ENCRYPTED_KEY
        );

        byte[] encryptedKey = new com.beingidly.litexl.crypto.AesCipher(keyDerivedKey)
            .encrypt(encryptionKey, encryptedKeySalt);

        // Generate and encrypt verifier
        byte[] verifierInput = new byte[16];
        random.nextBytes(verifierInput);

        byte[] verifierInputKey = com.beingidly.litexl.crypto.KeyDerivation.deriveKey(
            encryptionOptions.password(), encryptedKeySalt, encryptionOptions.spinCount(), keyBits,
            com.beingidly.litexl.crypto.KeyDerivation.BLOCK_KEY_VERIFIER_INPUT
        );
        byte[] encryptedVerifierInput = new com.beingidly.litexl.crypto.AesCipher(verifierInputKey)
            .encrypt(verifierInput, encryptedKeySalt);

        java.security.MessageDigest sha512 = java.security.MessageDigest.getInstance("SHA-512");
        byte[] verifierHash = sha512.digest(verifierInput);

        byte[] verifierHashKey = com.beingidly.litexl.crypto.KeyDerivation.deriveKey(
            encryptionOptions.password(), encryptedKeySalt, encryptionOptions.spinCount(), keyBits,
            com.beingidly.litexl.crypto.KeyDerivation.BLOCK_KEY_VERIFIER_VALUE
        );
        byte[] encryptedVerifierHash = new com.beingidly.litexl.crypto.AesCipher(verifierHashKey)
            .encryptNoPadding(verifierHash, encryptedKeySalt);

        // HMAC key is encrypted into EncryptionInfo; HMAC value itself is filled in after
        // EncryptedPackage is written, because it must cover the final stream bytes.
        byte[] hmacKey = new byte[64];
        random.nextBytes(hmacKey);

        com.beingidly.litexl.crypto.AesCipher hmacCipher = new com.beingidly.litexl.crypto.AesCipher(encryptionKey);

        byte[] hmacKeyIv = deriveHmacIv(keyDataSalt,
            com.beingidly.litexl.crypto.KeyDerivation.BLOCK_KEY_DATA_INTEGRITY_HMAC_KEY);
        byte[] encryptedHmacKey = hmacCipher.encryptNoPadding(hmacKey, hmacKeyIv);

        byte[] placeholderHmac = new byte[64];
        byte[] hmacValueIv = deriveHmacIv(keyDataSalt,
            com.beingidly.litexl.crypto.KeyDerivation.BLOCK_KEY_DATA_INTEGRITY_HMAC_VALUE);
        byte[] encryptedHmacValue = hmacCipher.encryptNoPadding(placeholderHmac, hmacValueIv);

        byte[] encryptionInfo = buildEncryptionInfo(
            keyBits, keyDataSalt, encryptedKeySalt, encryptionOptions.spinCount(),
            encryptedKey, encryptedVerifierInput, encryptedVerifierHash,
            encryptedHmacKey, encryptedHmacValue
        );

        return new EncryptionData(encryptionInfo, encryptionKey, keyDataSalt, hmacKey);
    }

    private record EncryptionData(
        byte[] encryptionInfo,
        byte[] encryptionKey,
        byte[] keyDataSalt,
        byte[] hmacKey
    ) {}

    private byte[] deriveHmacIv(byte[] salt, byte[] blockKey) throws GeneralSecurityException {
        java.security.MessageDigest sha512 = java.security.MessageDigest.getInstance("SHA-512");
        sha512.update(salt);
        sha512.update(blockKey);
        byte[] hash = sha512.digest();
        return java.util.Arrays.copyOf(hash, 16);
    }

    private byte[] buildEncryptionInfo(
            int keyBits, byte[] keyDataSalt, byte[] encryptedKeySalt, int spinCount,
            byte[] encryptedKey, byte[] encryptedVerifierInput, byte[] encryptedVerifierHash,
            byte[] encryptedHmacKey, byte[] encryptedHmacValue
    ) throws IOException {
        java.util.Base64.Encoder b64 = java.util.Base64.getEncoder();

        String xml = String.format("""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltSize="16" blockSize="16" keyBits="%d" hashSize="64"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       hashAlgorithm="SHA512" saltValue="%s"/>
              <dataIntegrity encryptedHmacKey="%s" encryptedHmacValue="%s"/>
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey spinCount="%d" saltSize="16" blockSize="16" keyBits="%d"
                                  hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  hashAlgorithm="SHA512" saltValue="%s"
                                  encryptedVerifierHashInput="%s"
                                  encryptedVerifierHashValue="%s"
                                  encryptedKeyValue="%s"/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
            """,
            keyBits, b64.encodeToString(keyDataSalt),
            b64.encodeToString(encryptedHmacKey), b64.encodeToString(encryptedHmacValue),
            spinCount, keyBits, b64.encodeToString(encryptedKeySalt),
            b64.encodeToString(encryptedVerifierInput), b64.encodeToString(encryptedVerifierHash),
            b64.encodeToString(encryptedKey)
        );

        byte[] xmlBytes = xml.getBytes(java.nio.charset.StandardCharsets.UTF_8);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        java.nio.ByteBuffer header = java.nio.ByteBuffer.allocate(8)
            .order(java.nio.ByteOrder.LITTLE_ENDIAN);
        header.putShort((short) 0x0004);
        header.putShort((short) 0x0004);
        header.putInt(0x00040);

        out.write(header.array());
        out.write(xmlBytes);

        return out.toByteArray();
    }

    @Override
    public void close() {
        // Nothing to close - ZipWriter is closed in write()
    }
}
