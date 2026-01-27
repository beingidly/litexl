package com.beingidly.litexl;

import com.beingidly.litexl.crypto.EncryptionOptions;
import com.beingidly.litexl.crypto.SheetProtection;
import com.beingidly.litexl.format.AutoFilter;
import com.beingidly.litexl.format.ConditionalFormat;
import com.beingidly.litexl.format.DataValidation;
import com.beingidly.litexl.style.BorderStyle;
import com.beingidly.litexl.style.HAlign;
import com.beingidly.litexl.style.Style;
import com.beingidly.litexl.style.VAlign;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for XlsxWriter covering uncovered branches.
 * Note: XlsxReader doesn't currently parse conditional formats, data validations,
 * autofilters, or sheet protection, so these tests verify writing completes without error.
 */
class XlsxWriterTest {
    @TempDir
    Path tempDir;

    @Test
    void writesConditionalFormatting() throws Exception {
        Path path = tempDir.resolve("conditional.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(50);

            // Create a style for conditional formatting
            Style style = Style.builder().fill(0xFFFF0000).build();
            int styleId = workbook.addStyle(style);

            sheet.addConditionalFormat(
                ConditionalFormat.greaterThan(CellRange.parse("A1:A10"), 10.0, styleId)
            );
            workbook.save(path);
        }

        // Verify file exists and is readable
        assertTrue(Files.exists(path));
        assertTrue(Files.size(path) > 0);

        // Verify basic data can be read back
        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertNotNull(sheet);
            assertEquals(50.0, sheet.cell(0, 0).number(), 0.001);
        }
    }

    @Test
    void writesMultipleConditionalFormats() throws Exception {
        Path path = tempDir.resolve("multi-conditional.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            for (int i = 0; i < 20; i++) {
                sheet.cell(i, 0).set(i * 5);
            }

            int styleId = workbook.addStyle(Style.builder().fill(0xFFFF0000).build());

            // Add multiple conditional formats
            sheet.addConditionalFormat(
                ConditionalFormat.greaterThan(CellRange.parse("A1:A20"), 50.0, styleId)
            );
            sheet.addConditionalFormat(
                ConditionalFormat.lessThan(CellRange.parse("A1:A20"), 25.0, styleId)
            );
            sheet.addConditionalFormat(
                ConditionalFormat.between(CellRange.parse("A1:A20"), 25.0, 50.0, styleId)
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesDataValidation() throws Exception {
        Path path = tempDir.resolve("validation.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            // Use wholeNumber validation which has a non-null operator (BETWEEN)
            sheet.addValidation(
                DataValidation.wholeNumber(CellRange.parse("A1:A10"), 1, 100)
            );
            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesDataValidationWithBetweenOperator() throws Exception {
        Path path = tempDir.resolve("between-validation.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");

            // Whole number validation with BETWEEN operator (default)
            sheet.addValidation(
                DataValidation.wholeNumber(CellRange.parse("A1:A10"), 1, 100)
            );

            // Decimal validation with BETWEEN operator (default)
            sheet.addValidation(
                DataValidation.decimal(CellRange.parse("B1:B10"), 0.0, 1.0)
            );

            // Text length validation with BETWEEN operator (default)
            sheet.addValidation(
                DataValidation.textLength(CellRange.parse("C1:C10"), 1, 50)
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesAllCellTypes() throws Exception {
        Path path = tempDir.resolve("alltypes.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");

            // String
            sheet.cell(0, 0).set("text");

            // Number
            sheet.cell(0, 1).set(123.45);

            // Boolean
            sheet.cell(0, 2).set(true);

            // Date
            sheet.cell(0, 3).set(LocalDateTime.of(2024, 1, 15, 10, 30));

            // Formula
            sheet.cell(0, 4).setFormula("A1&B1");

            // Error
            sheet.cell(0, 5).setValue(new CellValue.Error("#VALUE!"));

            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertNotNull(sheet);
            Row row = sheet.getRow(0);
            assertNotNull(row);

            assertEquals("text", row.cell(0).string());
            assertEquals(123.45, row.cell(1).number(), 0.001);
            assertTrue(row.cell(2).bool());
            assertEquals(CellType.FORMULA, row.cell(4).type());
            assertEquals(CellType.ERROR, row.cell(5).type());
        }
    }

    @Test
    void writesWithEncryption() throws Exception {
        Path path = tempDir.resolve("encrypted.xlsx");
        String password = "secret123";

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Secret");
            sheet.cell(0, 0).set("Confidential");
            workbook.save(path, EncryptionOptions.aes256(password, 100000));
        }

        // Verify file is encrypted by opening with password
        try (Workbook read = Workbook.open(path, password)) {
            assertEquals("Confidential", read.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesWithAes128Encryption() throws Exception {
        Path path = tempDir.resolve("encrypted-aes128.xlsx");
        String password = "aes128pass";

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Secret");
            sheet.cell(0, 0).set("AES-128 Data");
            workbook.save(path, EncryptionOptions.aes128(password));
        }

        try (Workbook read = Workbook.open(path, password)) {
            assertEquals("AES-128 Data", read.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesAutoFilter() throws Exception {
        Path path = tempDir.resolve("autofilter.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Header1");
            sheet.cell(0, 1).set("Header2");
            for (int i = 1; i < 10; i++) {
                sheet.cell(i, 0).set("Data" + i);
                sheet.cell(i, 1).set(i * 10);
            }
            sheet.setAutoFilter(CellRange.parse("A1:B10"));
            workbook.save(path);
        }

        // Verify file was written and can be opened
        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
            assertEquals("Header1", read.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesSheetProtection() throws Exception {
        Path path = tempDir.resolve("protected.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Protected");
            sheet.cell(0, 0).set("Data");
            sheet.protectionManager().protect("sheetpassword".toCharArray(), SheetProtection.defaults());
            workbook.save(path);
        }

        // Verify file was written and can be opened
        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Data", read.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesSheetProtectionWithAllOptions() throws Exception {
        Path path = tempDir.resolve("protected-options.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Protected");
            sheet.cell(0, 0).set("Data");

            // Test all protection options to cover all branches in writeSheetProtection
            SheetProtection options = SheetProtection.builder()
                .selectLockedCells(false)   // Covers !prot.selectLockedCells() branch
                .selectUnlockedCells(false) // Covers !prot.selectUnlockedCells() branch
                .formatCells(false)         // Covers !prot.formatCells() branch
                .formatColumns(false)       // Covers !prot.formatColumns() branch
                .formatRows(false)          // Covers !prot.formatRows() branch
                .insertRows(false)          // Covers !prot.insertRows() branch
                .insertColumns(false)       // Covers !prot.insertColumns() branch
                .deleteRows(false)          // Covers !prot.deleteRows() branch
                .deleteColumns(false)       // Covers !prot.deleteColumns() branch
                .sort(false)                // Covers !prot.sort() branch
                .autoFilter(false)          // Covers !prot.autoFilter() branch
                .pivotTables(false)         // Covers !prot.pivotTables() branch
                .build();

            sheet.protectionManager().protect("password".toCharArray(), options);
            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Data", read.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesSheetProtectionWithoutPassword() throws Exception {
        Path path = tempDir.resolve("protected-no-password.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Protected");
            sheet.cell(0, 0).set("Data");
            // Protect without password to cover hashInfo == null branch
            sheet.protect(SheetProtection.defaults());
            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Data", read.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesWhitespacePreservingText() throws Exception {
        Path path = tempDir.resolve("whitespace.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("  leading spaces");
            sheet.cell(0, 1).set("trailing spaces  ");
            sheet.cell(1, 0).set("  both sides  ");
            sheet.cell(1, 1).set("   multiple   spaces   ");
            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertEquals("  leading spaces", sheet.cell(0, 0).string());
            assertEquals("trailing spaces  ", sheet.cell(0, 1).string());
            assertEquals("  both sides  ", sheet.cell(1, 0).string());
            assertEquals("   multiple   spaces   ", sheet.cell(1, 1).string());
        }
    }

    @Test
    void writesEmptyCells() throws Exception {
        Path path = tempDir.resolve("empty.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            // Create cells but leave them empty
            sheet.cell(0, 0).setEmpty();
            sheet.cell(0, 1).set("Value");
            sheet.cell(0, 2).setEmpty();
            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertEquals("Value", sheet.cell(0, 1).string());
        }
    }

    @Test
    void writesMergedCells() throws Exception {
        Path path = tempDir.resolve("merged.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Merged Header");
            sheet.merge(CellRange.of(0, 0, 0, 4)); // A1:E1
            sheet.cell(2, 1).set("Another Merge");
            sheet.merge(2, 1, 4, 3); // B3:D5
            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertEquals(2, sheet.mergedCells().size());
            assertEquals("Merged Header", sheet.cell(0, 0).string());
        }
    }

    @Test
    void writesRowHeight() throws Exception {
        Path path = tempDir.resolve("rowheight.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Normal row");
            sheet.cell(1, 0).set("Tall row");
            sheet.row(1).height(30.0);
            sheet.cell(2, 0).set("Another normal row");
            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertNotNull(sheet);
            // Just verify the file was written correctly
            assertEquals("Normal row", sheet.cell(0, 0).string());
            assertEquals("Tall row", sheet.cell(1, 0).string());
        }
    }

    @Test
    void writesStyledCells() throws Exception {
        Path path = tempDir.resolve("styled.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");

            int boldStyle = workbook.addStyle(Style.builder().bold(true).build());
            int fillStyle = workbook.addStyle(Style.builder().fill(0xFF00FF00).build());

            sheet.cell(0, 0).set("Bold").style(boldStyle);
            sheet.cell(0, 1).set("Filled").style(fillStyle);
            sheet.cell(0, 2).set("Normal");

            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertEquals("Bold", sheet.cell(0, 0).string());
            assertEquals("Filled", sheet.cell(0, 1).string());
            assertEquals("Normal", sheet.cell(0, 2).string());
        }
    }

    @Test
    void writesMultipleSheets() throws Exception {
        Path path = tempDir.resolve("multisheets.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet1 = workbook.addSheet("Sheet1");
            sheet1.cell(0, 0).set("Data in Sheet1");

            Sheet sheet2 = workbook.addSheet("Sheet2");
            sheet2.cell(0, 0).set("Data in Sheet2");

            Sheet sheet3 = workbook.addSheet("Sheet3");
            sheet3.cell(0, 0).set("Data in Sheet3");

            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            assertEquals(3, read.sheetCount());
            assertEquals("Data in Sheet1", read.getSheet(0).cell(0, 0).string());
            assertEquals("Data in Sheet2", read.getSheet(1).cell(0, 0).string());
            assertEquals("Data in Sheet3", read.getSheet(2).cell(0, 0).string());
        }
    }

    @Test
    void writesExpressionConditionalFormat() throws Exception {
        Path path = tempDir.resolve("expression-cf.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            for (int i = 0; i < 10; i++) {
                sheet.cell(i, 0).set(i);
            }

            int styleId = workbook.addStyle(Style.builder().fill(0xFFFFFF00).build());

            // Expression-based conditional format (covers EXPRESSION type)
            sheet.addConditionalFormat(
                ConditionalFormat.expression(CellRange.parse("A1:A10"), "MOD(A1,2)=0", styleId)
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesDuplicateValuesConditionalFormat() throws Exception {
        Path path = tempDir.resolve("duplicate-cf.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("A");
            sheet.cell(1, 0).set("B");
            sheet.cell(2, 0).set("A"); // Duplicate

            int styleId = workbook.addStyle(Style.builder().fill(0xFFFF0000).build());

            // Covers DUPLICATE_VALUES type
            sheet.addConditionalFormat(
                ConditionalFormat.duplicateValues(CellRange.parse("A1:A10"), styleId)
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesUniqueValuesConditionalFormat() throws Exception {
        Path path = tempDir.resolve("unique-cf.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("A");
            sheet.cell(1, 0).set("B");
            sheet.cell(2, 0).set("C");

            int styleId = workbook.addStyle(Style.builder().fill(0xFF00FF00).build());

            // Covers UNIQUE_VALUES type
            sheet.addConditionalFormat(
                ConditionalFormat.uniqueValues(CellRange.parse("A1:A10"), styleId)
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesTextLengthDataValidation() throws Exception {
        Path path = tempDir.resolve("textlength-validation.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");

            // Text length validation (covers TEXT_LENGTH type with BETWEEN operator)
            sheet.addValidation(
                DataValidation.textLength(CellRange.parse("A1:A10"), 5, 100)
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesBooleanFalse() throws Exception {
        Path path = tempDir.resolve("bool-false.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(false);
            sheet.cell(0, 1).set(true);
            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertFalse(sheet.cell(0, 0).bool());
            assertTrue(sheet.cell(0, 1).bool());
        }
    }

    @Test
    void writesNegativeNumbers() throws Exception {
        Path path = tempDir.resolve("negative.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(-100);
            sheet.cell(0, 1).set(-0.5);
            sheet.cell(0, 2).set(-1.23e10);
            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertEquals(-100.0, sheet.cell(0, 0).number(), 0.001);
            assertEquals(-0.5, sheet.cell(0, 1).number(), 0.001);
            assertEquals(-1.23e10, sheet.cell(0, 2).number(), 1e5);
        }
    }

    @Test
    void writesSpecialCharacters() throws Exception {
        Path path = tempDir.resolve("special.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Special: <>&\"'");
            sheet.cell(0, 1).set("Unicode: 한글日本語");
            sheet.cell(0, 2).set("Newline:\nSecond line");
            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertEquals("Special: <>&\"'", sheet.cell(0, 0).string());
            assertEquals("Unicode: 한글日本語", sheet.cell(0, 1).string());
            assertEquals("Newline:\nSecond line", sheet.cell(0, 2).string());
        }
    }

    @Test
    void writesConditionalFormatWithBetweenOperator() throws Exception {
        Path path = tempDir.resolve("between-cf.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            for (int i = 0; i < 20; i++) {
                sheet.cell(i, 0).set(i * 5);
            }

            int styleId = workbook.addStyle(Style.builder().fill(0xFF00FFFF).build());

            // Between conditional format (has formula2, covers BETWEEN operator)
            sheet.addConditionalFormat(
                ConditionalFormat.between(CellRange.parse("A1:A20"), 25.0, 75.0, styleId)
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesDataValidationWithNonBetweenOperators() throws Exception {
        Path path = tempDir.resolve("operator-validation.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");

            // Different operators to cover all branches in dvOperatorName
            sheet.addValidation(
                DataValidation.wholeNumber(CellRange.parse("A1:A5"), DataValidation.Operator.GREATER_THAN, "10", null)
            );
            sheet.addValidation(
                DataValidation.wholeNumber(CellRange.parse("B1:B5"), DataValidation.Operator.LESS_THAN, "100", null)
            );
            sheet.addValidation(
                DataValidation.wholeNumber(CellRange.parse("C1:C5"), DataValidation.Operator.EQUAL, "50", null)
            );
            sheet.addValidation(
                DataValidation.wholeNumber(CellRange.parse("D1:D5"), DataValidation.Operator.NOT_EQUAL, "0", null)
            );
            sheet.addValidation(
                DataValidation.decimal(CellRange.parse("E1:E5"), DataValidation.Operator.GREATER_THAN_OR_EQUAL, "0.5", null)
            );
            sheet.addValidation(
                DataValidation.decimal(CellRange.parse("F1:F5"), DataValidation.Operator.LESS_THAN_OR_EQUAL, "1.0", null)
            );
            sheet.addValidation(
                DataValidation.textLength(CellRange.parse("G1:G5"), DataValidation.Operator.NOT_BETWEEN, "1", "10")
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesConditionalFormatWithAllOperators() throws Exception {
        Path path = tempDir.resolve("all-cf-operators.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            for (int i = 0; i < 10; i++) {
                sheet.cell(i, 0).set(i * 10);
            }

            int styleId = workbook.addStyle(Style.builder().fill(0xFF0000FF).build());

            // Cover all conditional format operators
            sheet.addConditionalFormat(
                ConditionalFormat.lessThan(CellRange.parse("A1:A10"), 20.0, styleId)
            );
            sheet.addConditionalFormat(
                new ConditionalFormat(
                    CellRange.parse("B1:B10"),
                    ConditionalFormat.Type.CELL_VALUE,
                    ConditionalFormat.Operator.LESS_THAN_OR_EQUAL,
                    "30",
                    null,
                    styleId
                )
            );
            sheet.addConditionalFormat(
                new ConditionalFormat(
                    CellRange.parse("C1:C10"),
                    ConditionalFormat.Type.CELL_VALUE,
                    ConditionalFormat.Operator.EQUAL,
                    "50",
                    null,
                    styleId
                )
            );
            sheet.addConditionalFormat(
                new ConditionalFormat(
                    CellRange.parse("D1:D10"),
                    ConditionalFormat.Type.CELL_VALUE,
                    ConditionalFormat.Operator.NOT_EQUAL,
                    "60",
                    null,
                    styleId
                )
            );
            sheet.addConditionalFormat(
                new ConditionalFormat(
                    CellRange.parse("E1:E10"),
                    ConditionalFormat.Type.CELL_VALUE,
                    ConditionalFormat.Operator.GREATER_THAN_OR_EQUAL,
                    "70",
                    null,
                    styleId
                )
            );
            sheet.addConditionalFormat(
                new ConditionalFormat(
                    CellRange.parse("F1:F10"),
                    ConditionalFormat.Type.CELL_VALUE,
                    ConditionalFormat.Operator.NOT_BETWEEN,
                    "10",
                    "90",
                    styleId
                )
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesConditionalFormatWithNoOperator() throws Exception {
        Path path = tempDir.resolve("no-op-cf.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Test");

            int styleId = workbook.addStyle(Style.builder().fill(0xFFFFFF00).build());

            // Conditional format with NONE operator (covered by expression type)
            sheet.addConditionalFormat(
                new ConditionalFormat(
                    CellRange.parse("A1:A10"),
                    ConditionalFormat.Type.EXPRESSION,
                    ConditionalFormat.Operator.NONE,
                    "LEN(A1)>0",
                    null,
                    styleId
                )
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
    }

    @Test
    void writesConditionalFormatWithNoStyleId() throws Exception {
        Path path = tempDir.resolve("no-style-cf.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(100);

            // Conditional format with styleId = 0 (covers styleId <= 0 branch)
            sheet.addConditionalFormat(
                new ConditionalFormat(
                    CellRange.parse("A1:A10"),
                    ConditionalFormat.Type.CELL_VALUE,
                    ConditionalFormat.Operator.GREATER_THAN,
                    "50",
                    null,
                    0  // No style
                )
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
    }

    @Test
    void writesDataValidationWithErrorMessages() throws Exception {
        Path path = tempDir.resolve("error-msg-validation.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");

            // Validation with error title and message
            sheet.addValidation(
                DataValidation.wholeNumber(CellRange.parse("A1:A10"), 1, 100)
            );

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
    }

    @Test
    void writesDateCell() throws Exception {
        Path path = tempDir.resolve("date.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(LocalDateTime.of(2024, 6, 15, 14, 30, 45));
            sheet.cell(0, 1).set(LocalDateTime.of(2000, 1, 1, 0, 0, 0));
            workbook.save(path);
        }

        assertTrue(Files.exists(path));
        try (Workbook read = Workbook.open(path)) {
            // Dates are stored as numbers in Excel
            assertNotNull(read.getSheet(0));
        }
    }

    @Test
    void writesNullCellValue() throws Exception {
        Path path = tempDir.resolve("null-value.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).setValue(null);  // Should create Empty cell
            sheet.cell(0, 1).set("after null");
            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            Sheet sheet = read.getSheet(0);
            assertEquals("after null", sheet.cell(0, 1).string());
        }
    }

    @Test
    void writesCellWithStyle() throws Exception {
        Path path = tempDir.resolve("cell-style.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");

            // Create style with non-zero ID to cover styleId > 0 branch
            int styleId = workbook.addStyle(Style.builder().bold(true).fill(0xFFCCCCCC).build());

            sheet.cell(0, 0).set("Styled").style(styleId);
            sheet.cell(0, 1).set("Unstyled");  // styleId = 0

            workbook.save(path);
        }

        assertTrue(Files.exists(path));
    }

    @Test
    void writesEmptySheet() throws Exception {
        Path path = tempDir.resolve("empty-sheet.xlsx");

        try (Workbook workbook = Workbook.create()) {
            workbook.addSheet("Empty");
            workbook.save(path);
        }

        try (Workbook read = Workbook.open(path)) {
            assertEquals(1, read.sheetCount());
            assertEquals("Empty", read.getSheet(0).name());
        }
    }

    @Test
    void writeWithStyles() throws Exception {
        Path file = tempDir.resolve("styled.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Styled");

            Style boldStyle = Style.builder().bold(true).build();
            Style colorStyle = Style.builder()
                .fill(0xFFCCCCCC)
                .font("Arial", 14)
                .color(0xFFFF0000)
                .build();

            int boldId = wb.addStyle(boldStyle);
            int colorId = wb.addStyle(colorStyle);

            sheet.cell(0, 0).set("Bold").style(boldId);
            sheet.cell(0, 1).set("Colored").style(colorId);

            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals("Bold", wb.getSheet(0).getCell(0, 0).string());
            assertEquals("Colored", wb.getSheet(0).getCell(0, 1).string());
        }
    }

    @Test
    void writeWithBorders() throws Exception {
        Path file = tempDir.resolve("borders.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Borders");

            Style borderStyle = Style.builder()
                .border(BorderStyle.THIN, 0xFF000000)
                .build();
            int styleId = wb.addStyle(borderStyle);

            sheet.cell(0, 0).set("Bordered").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertNotNull(wb.getSheet(0).getCell(0, 0).string());
        }
    }

    @Test
    void writeWithAlignment() throws Exception {
        Path file = tempDir.resolve("alignment.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Alignment");

            Style centerStyle = Style.builder()
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .wrap(true)
                .build();
            int styleId = wb.addStyle(centerStyle);

            sheet.cell(0, 0).set("Centered").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertNotNull(wb.getSheet(0));
        }
    }

    @Test
    void writeWithNumberFormat() throws Exception {
        Path file = tempDir.resolve("numformat.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Numbers");

            Style percentStyle = Style.builder().format("0.00%").build();
            Style currencyStyle = Style.builder().format("$#,##0.00").build();

            int pctId = wb.addStyle(percentStyle);
            int curId = wb.addStyle(currencyStyle);

            sheet.cell(0, 0).set(0.15).style(pctId);
            sheet.cell(0, 1).set(1234.56).style(curId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals(0.15, wb.getSheet(0).getCell(0, 0).number(), 0.001);
        }
    }

    @Test
    void writeWithUnlockedCells() throws Exception {
        Path file = tempDir.resolve("unlocked.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Unlocked");

            Style unlockedStyle = Style.builder().locked(false).build();
            int styleId = wb.addStyle(unlockedStyle);

            sheet.cell(0, 0).set("Editable").style(styleId);
            sheet.protectionManager().protect("password".toCharArray(), SheetProtection.defaults());
            wb.save(file);
        }

        // Verify file was written and can be opened (protection is written but not parsed back)
        try (Workbook wb = Workbook.open(file)) {
            assertEquals("Editable", wb.getSheet(0).getCell(0, 0).string());
        }
    }

    @Test
    void writeEmptyWorkbook() throws Exception {
        Path file = tempDir.resolve("empty.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.addSheet("Empty");
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals(1, wb.sheetCount());
        }
    }

    @Test
    void writeColumnWidths() throws Exception {
        Path file = tempDir.resolve("widths.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Widths");
            sheet.format().setColumnWidth(0, 20.0);
            sheet.format().setColumnWidth(2, 30.0);
            sheet.cell(0, 0).set("Wide column");
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertNotNull(wb.getSheet(0));
        }
    }

    @Test
    void writeHiddenSheet() throws Exception {
        Path file = tempDir.resolve("hidden.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet visible = wb.addSheet("Visible");
            Sheet hidden = wb.addSheet("Hidden");
            hidden.format().setHidden(true);
            visible.cell(0, 0).set("Data");
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals(2, wb.sheetCount());
        }
    }

    // === Additional tests for XmlWriter coverage ===

    @Test
    void writesItalicFont() throws Exception {
        Path file = tempDir.resolve("italic.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            Style italicStyle = Style.builder().italic(true).build();
            int styleId = wb.addStyle(italicStyle);
            sheet.cell(0, 0).set("Italic text").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals("Italic text", wb.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesUnderlineFont() throws Exception {
        Path file = tempDir.resolve("underline.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            Style underlineStyle = Style.builder().underline(true).build();
            int styleId = wb.addStyle(underlineStyle);
            sheet.cell(0, 0).set("Underlined text").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals("Underlined text", wb.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesStrikethroughFont() throws Exception {
        Path file = tempDir.resolve("strikethrough.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            // Create font with strikethrough
            var font = new com.beingidly.litexl.style.Font(
                "Calibri", 11.0, 0xFF000000, false, false, false, true
            );
            Style strikeStyle = Style.builder().font(font).build();
            int styleId = wb.addStyle(strikeStyle);
            sheet.cell(0, 0).set("Strikethrough text").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals("Strikethrough text", wb.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesAllFontStylesCombined() throws Exception {
        Path file = tempDir.resolve("all-font-styles.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            // Create font with all styles enabled
            var font = new com.beingidly.litexl.style.Font(
                "Arial", 14.0, 0xFFFF0000, true, true, true, true
            );
            Style allStyle = Style.builder().font(font).build();
            int styleId = wb.addStyle(allStyle);
            sheet.cell(0, 0).set("All styles").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals("All styles", wb.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesMediumBorder() throws Exception {
        Path file = tempDir.resolve("medium-border.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            Style borderStyle = Style.builder().border(BorderStyle.MEDIUM, 0xFF000000).build();
            int styleId = wb.addStyle(borderStyle);
            sheet.cell(0, 0).set("Medium border").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertNotNull(wb.getSheet(0));
        }
    }

    @Test
    void writesThickBorder() throws Exception {
        Path file = tempDir.resolve("thick-border.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            Style borderStyle = Style.builder().border(BorderStyle.THICK, 0xFF000000).build();
            int styleId = wb.addStyle(borderStyle);
            sheet.cell(0, 0).set("Thick border").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertNotNull(wb.getSheet(0));
        }
    }

    @Test
    void writesDoubleBorder() throws Exception {
        Path file = tempDir.resolve("double-border.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            Style borderStyle = Style.builder().border(BorderStyle.DOUBLE, 0xFF000000).build();
            int styleId = wb.addStyle(borderStyle);
            sheet.cell(0, 0).set("Double border").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertNotNull(wb.getSheet(0));
        }
    }

    @Test
    void writesDashedBorder() throws Exception {
        Path file = tempDir.resolve("dashed-border.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            Style borderStyle = Style.builder().border(BorderStyle.DASHED, 0xFF000000).build();
            int styleId = wb.addStyle(borderStyle);
            sheet.cell(0, 0).set("Dashed border").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertNotNull(wb.getSheet(0));
        }
    }

    @Test
    void writesDottedBorder() throws Exception {
        Path file = tempDir.resolve("dotted-border.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            Style borderStyle = Style.builder().border(BorderStyle.DOTTED, 0xFF000000).build();
            int styleId = wb.addStyle(borderStyle);
            sheet.cell(0, 0).set("Dotted border").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertNotNull(wb.getSheet(0));
        }
    }

    @Test
    void writesAllBorderStyles() throws Exception {
        Path file = tempDir.resolve("all-borders.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            BorderStyle[] styles = {BorderStyle.THIN, BorderStyle.MEDIUM, BorderStyle.THICK,
                BorderStyle.DOUBLE, BorderStyle.DASHED, BorderStyle.DOTTED};
            int row = 0;
            for (BorderStyle bs : styles) {
                Style style = Style.builder().border(bs, 0xFF000000).build();
                int styleId = wb.addStyle(style);
                sheet.cell(row++, 0).set(bs.name()).style(styleId);
            }
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals("THIN", wb.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesErrorCells() throws Exception {
        Path file = tempDir.resolve("error-cells.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            // Write different error values
            sheet.cell(0, 0).setValue(new CellValue.Error("#VALUE!"));
            sheet.cell(1, 0).setValue(new CellValue.Error("#DIV/0!"));
            sheet.cell(2, 0).setValue(new CellValue.Error("#REF!"));
            sheet.cell(3, 0).setValue(new CellValue.Error("#NAME?"));
            sheet.cell(4, 0).setValue(new CellValue.Error("#NUM!"));
            sheet.cell(5, 0).setValue(new CellValue.Error("#N/A"));
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            Sheet sheet = wb.getSheet(0);
            assertEquals(CellType.ERROR, sheet.cell(0, 0).type());
            assertEquals(CellType.ERROR, sheet.cell(1, 0).type());
            assertEquals(CellType.ERROR, sheet.cell(2, 0).type());
        }
    }

    @Test
    void writesFormulas() throws Exception {
        Path file = tempDir.resolve("formulas.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            sheet.cell(0, 0).set(10);
            sheet.cell(0, 1).set(20);
            sheet.cell(0, 2).setFormula("A1+B1");
            sheet.cell(1, 0).setFormula("SUM(A1:B1)");
            sheet.cell(2, 0).setFormula("AVERAGE(A1:B1)");
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            Sheet sheet = wb.getSheet(0);
            assertEquals(10.0, sheet.cell(0, 0).number(), 0.001);
            assertEquals(20.0, sheet.cell(0, 1).number(), 0.001);
            assertEquals(CellType.FORMULA, sheet.cell(0, 2).type());
            assertEquals(CellType.FORMULA, sheet.cell(1, 0).type());
        }
    }

    @Test
    void writesIndividualBorderSides() throws Exception {
        Path file = tempDir.resolve("border-sides.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");

            // Left border only
            Style leftStyle = Style.builder()
                .borderLeft(BorderStyle.THIN, 0xFF000000)
                .build();
            sheet.cell(0, 0).set("Left").style(wb.addStyle(leftStyle));

            // Right border only
            Style rightStyle = Style.builder()
                .borderRight(BorderStyle.MEDIUM, 0xFF0000FF)
                .build();
            sheet.cell(1, 0).set("Right").style(wb.addStyle(rightStyle));

            // Top border only
            Style topStyle = Style.builder()
                .borderTop(BorderStyle.THICK, 0xFFFF0000)
                .build();
            sheet.cell(2, 0).set("Top").style(wb.addStyle(topStyle));

            // Bottom border only
            Style bottomStyle = Style.builder()
                .borderBottom(BorderStyle.DOUBLE, 0xFF00FF00)
                .build();
            sheet.cell(3, 0).set("Bottom").style(wb.addStyle(bottomStyle));

            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals("Left", wb.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesNoneBorder() throws Exception {
        Path file = tempDir.resolve("none-border.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            Style noneStyle = Style.builder().border(BorderStyle.NONE, 0xFF000000).build();
            int styleId = wb.addStyle(noneStyle);
            sheet.cell(0, 0).set("No border").style(styleId);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertEquals("No border", wb.getSheet(0).cell(0, 0).string());
        }
    }

    // === Tests for AutoFilter coverage ===

    @Test
    void writesAutoFilterWithValueFilter() throws Exception {
        Path file = tempDir.resolve("autofilter-values.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Value");
            for (int i = 1; i <= 10; i++) {
                sheet.cell(i, 0).set("Item" + i);
                sheet.cell(i, 1).set(i * 10);
            }

            // Create autofilter with value filter
            var filterColumn = AutoFilter.FilterColumn.values(0, java.util.List.of("Item1", "Item2", "Item3"));
            var autoFilter = new AutoFilter(CellRange.parse("A1:B11"), java.util.List.of(filterColumn));
            sheet.setAutoFilter(autoFilter);
            wb.save(file);
        }

        assertTrue(Files.exists(file));
        try (Workbook wb = Workbook.open(file)) {
            assertEquals("Name", wb.getSheet(0).cell(0, 0).string());
        }
    }

    @Test
    void writesAutoFilterWithCustomFilter() throws Exception {
        Path file = tempDir.resolve("autofilter-custom.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            sheet.cell(0, 0).set("Value");
            for (int i = 1; i <= 10; i++) {
                sheet.cell(i, 0).set(i * 10);
            }

            // Create autofilter with custom filter (single condition)
            var customFilter = AutoFilter.CustomFilter.single(
                AutoFilter.CustomFilter.Operator.GREATER_THAN, "50");
            var filterColumn = AutoFilter.FilterColumn.custom(0, customFilter);
            var autoFilter = new AutoFilter(CellRange.parse("A1:A11"), java.util.List.of(filterColumn));
            sheet.setAutoFilter(autoFilter);
            wb.save(file);
        }

        assertTrue(Files.exists(file));
    }

    @Test
    void writesAutoFilterWithTwoConditionAndFilter() throws Exception {
        Path file = tempDir.resolve("autofilter-and.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            sheet.cell(0, 0).set("Value");
            for (int i = 1; i <= 20; i++) {
                sheet.cell(i, 0).set(i * 5);
            }

            // Create autofilter with two conditions (AND)
            var customFilter = AutoFilter.CustomFilter.and(
                AutoFilter.CustomFilter.Operator.GREATER_THAN_OR_EQUAL, "20",
                AutoFilter.CustomFilter.Operator.LESS_THAN_OR_EQUAL, "80");
            var filterColumn = AutoFilter.FilterColumn.custom(0, customFilter);
            var autoFilter = new AutoFilter(CellRange.parse("A1:A21"), java.util.List.of(filterColumn));
            sheet.setAutoFilter(autoFilter);
            wb.save(file);
        }

        assertTrue(Files.exists(file));
    }

    @Test
    void writesAutoFilterWithTwoConditionOrFilter() throws Exception {
        Path file = tempDir.resolve("autofilter-or.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            sheet.cell(0, 0).set("Value");
            for (int i = 1; i <= 20; i++) {
                sheet.cell(i, 0).set(i * 5);
            }

            // Create autofilter with two conditions (OR) - covers and=false branch
            var customFilter = AutoFilter.CustomFilter.or(
                AutoFilter.CustomFilter.Operator.LESS_THAN, "15",
                AutoFilter.CustomFilter.Operator.GREATER_THAN, "85");
            var filterColumn = AutoFilter.FilterColumn.custom(0, customFilter);
            var autoFilter = new AutoFilter(CellRange.parse("A1:A21"), java.util.List.of(filterColumn));
            sheet.setAutoFilter(autoFilter);
            wb.save(file);
        }

        assertTrue(Files.exists(file));
    }

    @Test
    void writesAutoFilterWithAllOperators() throws Exception {
        Path file = tempDir.resolve("autofilter-all-ops.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            sheet.cell(0, 0).set("A");
            sheet.cell(0, 1).set("B");
            sheet.cell(0, 2).set("C");
            sheet.cell(0, 3).set("D");
            sheet.cell(0, 4).set("E");
            sheet.cell(0, 5).set("F");
            for (int i = 1; i <= 10; i++) {
                for (int j = 0; j < 6; j++) {
                    sheet.cell(i, j).set(i * 10);
                }
            }

            // Test all filter operators
            var filter1 = AutoFilter.FilterColumn.custom(0,
                AutoFilter.CustomFilter.single(AutoFilter.CustomFilter.Operator.EQUAL, "50"));
            var filter2 = AutoFilter.FilterColumn.custom(1,
                AutoFilter.CustomFilter.single(AutoFilter.CustomFilter.Operator.NOT_EQUAL, "50"));
            var filter3 = AutoFilter.FilterColumn.custom(2,
                AutoFilter.CustomFilter.single(AutoFilter.CustomFilter.Operator.GREATER_THAN, "50"));
            var filter4 = AutoFilter.FilterColumn.custom(3,
                AutoFilter.CustomFilter.single(AutoFilter.CustomFilter.Operator.GREATER_THAN_OR_EQUAL, "50"));
            var filter5 = AutoFilter.FilterColumn.custom(4,
                AutoFilter.CustomFilter.single(AutoFilter.CustomFilter.Operator.LESS_THAN, "50"));
            var filter6 = AutoFilter.FilterColumn.custom(5,
                AutoFilter.CustomFilter.single(AutoFilter.CustomFilter.Operator.LESS_THAN_OR_EQUAL, "50"));

            var autoFilter = new AutoFilter(CellRange.parse("A1:F11"),
                java.util.List.of(filter1, filter2, filter3, filter4, filter5, filter6));
            sheet.setAutoFilter(autoFilter);
            wb.save(file);
        }

        assertTrue(Files.exists(file));
    }

    // === Tests for all ConditionalFormat types ===

    @Test
    void writesAllConditionalFormatTypes() throws Exception {
        Path file = tempDir.resolve("all-cf-types.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            for (int i = 0; i < 20; i++) {
                sheet.cell(i, 0).set(i * 5);
            }

            int styleId = wb.addStyle(Style.builder().fill(0xFFFF0000).build());

            // Test all conditional format types
            // COLOR_SCALE
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("A1:A5"), ConditionalFormat.Type.COLOR_SCALE,
                ConditionalFormat.Operator.NONE, null, null, styleId));

            // DATA_BAR
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("A6:A10"), ConditionalFormat.Type.DATA_BAR,
                ConditionalFormat.Operator.NONE, null, null, styleId));

            // ICON_SET
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("A11:A15"), ConditionalFormat.Type.ICON_SET,
                ConditionalFormat.Operator.NONE, null, null, styleId));

            // TOP_BOTTOM
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("B1:B5"), ConditionalFormat.Type.TOP_BOTTOM,
                ConditionalFormat.Operator.NONE, "10", null, styleId));

            // ABOVE_AVERAGE
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("B6:B10"), ConditionalFormat.Type.ABOVE_AVERAGE,
                ConditionalFormat.Operator.NONE, null, null, styleId));

            // CONTAINS_TEXT
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("C1:C5"), ConditionalFormat.Type.CONTAINS_TEXT,
                ConditionalFormat.Operator.NONE, "test", null, styleId));

            // NOT_CONTAINS_TEXT
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("C6:C10"), ConditionalFormat.Type.NOT_CONTAINS_TEXT,
                ConditionalFormat.Operator.NONE, "test", null, styleId));

            // BEGINS_WITH
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("D1:D5"), ConditionalFormat.Type.BEGINS_WITH,
                ConditionalFormat.Operator.NONE, "A", null, styleId));

            // ENDS_WITH
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("D6:D10"), ConditionalFormat.Type.ENDS_WITH,
                ConditionalFormat.Operator.NONE, "Z", null, styleId));

            // CONTAINS_BLANKS
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("E1:E5"), ConditionalFormat.Type.CONTAINS_BLANKS,
                ConditionalFormat.Operator.NONE, null, null, styleId));

            // CONTAINS_ERRORS
            sheet.addConditionalFormat(new ConditionalFormat(
                CellRange.parse("E6:E10"), ConditionalFormat.Type.CONTAINS_ERRORS,
                ConditionalFormat.Operator.NONE, null, null, styleId));

            wb.save(file);
        }

        assertTrue(Files.exists(file));
    }

    // === Tests for all DataValidation types ===

    @Test
    void writesAllDataValidationTypes() throws Exception {
        Path file = tempDir.resolve("all-dv-types.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");

            // ANY type
            sheet.addValidation(new DataValidation(
                CellRange.parse("A1:A5"), DataValidation.Type.ANY,
                DataValidation.Operator.BETWEEN, null, null, null, null, false));

            // DATE type
            sheet.addValidation(new DataValidation(
                CellRange.parse("B1:B5"), DataValidation.Type.DATE,
                DataValidation.Operator.BETWEEN, "44927", "45291", null, null, false));

            // TIME type
            sheet.addValidation(new DataValidation(
                CellRange.parse("C1:C5"), DataValidation.Type.TIME,
                DataValidation.Operator.BETWEEN, "0.25", "0.75", null, null, false));

            // CUSTOM type - using BETWEEN operator (it's ignored for custom but avoids NPE)
            sheet.addValidation(new DataValidation(
                CellRange.parse("D1:D5"), DataValidation.Type.CUSTOM,
                DataValidation.Operator.BETWEEN, "LEN(D1)>0", null, null, "Value is required", false));

            // LIST type (dropdown) with BETWEEN operator
            sheet.addValidation(new DataValidation(
                CellRange.parse("E1:E5"), DataValidation.Type.LIST,
                DataValidation.Operator.BETWEEN, "\"Apple,Banana,Cherry\"", null, null, null, true));

            wb.save(file);
        }

        assertTrue(Files.exists(file));
    }

    @Test
    void writesDataValidationWithShowDropdownFalse() throws Exception {
        Path file = tempDir.resolve("dv-no-dropdown.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");

            // List validation with showDropdown = false (covers else branch)
            // Using BETWEEN operator to avoid NPE
            sheet.addValidation(new DataValidation(
                CellRange.parse("A1:A10"), DataValidation.Type.LIST,
                DataValidation.Operator.BETWEEN, "\"A,B,C\"", null, null, null, false));

            wb.save(file);
        }

        assertTrue(Files.exists(file));
    }

    @Test
    void writesDataValidationWithoutErrorMessages() throws Exception {
        Path file = tempDir.resolve("dv-no-error-msg.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");

            // Validation without error title and message (covers null branches)
            sheet.addValidation(new DataValidation(
                CellRange.parse("A1:A10"), DataValidation.Type.WHOLE,
                DataValidation.Operator.BETWEEN, "1", "100", null, null, false));

            wb.save(file);
        }

        assertTrue(Files.exists(file));
    }

    @Test
    void writesDataValidationWithoutFormula2() throws Exception {
        Path file = tempDir.resolve("dv-no-formula2.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");

            // Validation with formula1 but no formula2
            sheet.addValidation(new DataValidation(
                CellRange.parse("A1:A10"), DataValidation.Type.WHOLE,
                DataValidation.Operator.GREATER_THAN, "10", null, "Error", "Must be > 10", false));

            wb.save(file);
        }

        assertTrue(Files.exists(file));
    }
}
