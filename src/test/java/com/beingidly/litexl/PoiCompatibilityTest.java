package com.beingidly.litexl;

import com.beingidly.litexl.crypto.EncryptionOptions;
import com.beingidly.litexl.crypto.SheetProtection;
import com.beingidly.litexl.style.*;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.DisabledInNativeImage;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests that verify litexl-generated XLSX files can be read correctly by Apache POI.
 * This ensures compatibility with other Excel-reading applications.
 */
@DisabledInNativeImage
class PoiCompatibilityTest {

    @TempDir
    Path tempDir;

    @Test
    void stringCells() throws Exception {
        Path file = tempDir.resolve("strings.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Strings");
            sheet.cell(0, 0).set("Hello");
            sheet.cell(0, 1).set("World");
            sheet.cell(1, 0).set("한글 테스트");
            sheet.cell(1, 1).set("Special: <>&\"'");
            sheet.cell(2, 0).set("  spaces  ");
            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);
            assertEquals("Strings", sheet.getSheetName());

            assertEquals("Hello", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("World", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals("한글 테스트", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("Special: <>&\"'", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("  spaces  ", sheet.getRow(2).getCell(0).getStringCellValue());
        }
    }

    @Test
    void numericCells() throws Exception {
        Path file = tempDir.resolve("numbers.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Numbers");
            sheet.cell(0, 0).set(0);
            sheet.cell(0, 1).set(42);
            sheet.cell(0, 2).set(-100);
            sheet.cell(1, 0).set(3.14159);
            sheet.cell(1, 1).set(1.23e10);
            sheet.cell(1, 2).set(-0.001);
            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);

            assertEquals(0.0, sheet.getRow(0).getCell(0).getNumericCellValue());
            assertEquals(42.0, sheet.getRow(0).getCell(1).getNumericCellValue());
            assertEquals(-100.0, sheet.getRow(0).getCell(2).getNumericCellValue());
            assertEquals(3.14159, sheet.getRow(1).getCell(0).getNumericCellValue(), 0.00001);
            assertEquals(1.23e10, sheet.getRow(1).getCell(1).getNumericCellValue(), 1e5);
            assertEquals(-0.001, sheet.getRow(1).getCell(2).getNumericCellValue(), 0.0001);
        }
    }

    @Test
    void booleanCells() throws Exception {
        Path file = tempDir.resolve("booleans.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Booleans");
            sheet.cell(0, 0).set(true);
            sheet.cell(0, 1).set(false);
            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);

            assertEquals(CellType.BOOLEAN, sheet.getRow(0).getCell(0).getCellType());
            assertTrue(sheet.getRow(0).getCell(0).getBooleanCellValue());
            assertFalse(sheet.getRow(0).getCell(1).getBooleanCellValue());
        }
    }

    @Test
    void formulaCells() throws Exception {
        Path file = tempDir.resolve("formulas.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Formulas");
            sheet.cell(0, 0).set(10);
            sheet.cell(0, 1).set(20);
            sheet.cell(0, 2).setFormula("A1+B1");
            sheet.cell(1, 0).setFormula("SUM(A1:B1)");
            sheet.cell(1, 1).setFormula("AVERAGE(A1:B1)");
            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);

            assertEquals(CellType.FORMULA, sheet.getRow(0).getCell(2).getCellType());
            assertEquals("A1+B1", sheet.getRow(0).getCell(2).getCellFormula());
            assertEquals("SUM(A1:B1)", sheet.getRow(1).getCell(0).getCellFormula());
            assertEquals("AVERAGE(A1:B1)", sheet.getRow(1).getCell(1).getCellFormula());
        }
    }

    @Test
    void mixedTypes() throws Exception {
        Path file = tempDir.resolve("mixed.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Mixed");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Value");
            sheet.cell(0, 2).set("Active");

            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(100);
            sheet.cell(1, 2).set(true);

            sheet.cell(2, 0).set("Bob");
            sheet.cell(2, 1).set(200);
            sheet.cell(2, 2).set(false);

            sheet.cell(3, 0).set("Total");
            sheet.cell(3, 1).setFormula("SUM(B2:B3)");
            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);

            // Headers
            assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Value", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals("Active", sheet.getRow(0).getCell(2).getStringCellValue());

            // Data rows
            assertEquals("Alice", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(100.0, sheet.getRow(1).getCell(1).getNumericCellValue());
            assertTrue(sheet.getRow(1).getCell(2).getBooleanCellValue());

            assertEquals("Bob", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals(200.0, sheet.getRow(2).getCell(1).getNumericCellValue());
            assertFalse(sheet.getRow(2).getCell(2).getBooleanCellValue());

            // Formula row
            assertEquals("Total", sheet.getRow(3).getCell(0).getStringCellValue());
            assertEquals("SUM(B2:B3)", sheet.getRow(3).getCell(1).getCellFormula());
        }
    }

    @Test
    void multipleSheets() throws Exception {
        Path file = tempDir.resolve("multi-sheet.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet1 = wb.addSheet("First");
            sheet1.cell(0, 0).set("Sheet 1 Data");

            Sheet sheet2 = wb.addSheet("Second");
            sheet2.cell(0, 0).set("Sheet 2 Data");

            Sheet sheet3 = wb.addSheet("Third");
            sheet3.cell(0, 0).set("Sheet 3 Data");

            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            assertEquals(3, wb.getNumberOfSheets());
            assertEquals("First", wb.getSheetAt(0).getSheetName());
            assertEquals("Second", wb.getSheetAt(1).getSheetName());
            assertEquals("Third", wb.getSheetAt(2).getSheetName());

            assertEquals("Sheet 1 Data", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertEquals("Sheet 2 Data", wb.getSheetAt(1).getRow(0).getCell(0).getStringCellValue());
            assertEquals("Sheet 3 Data", wb.getSheetAt(2).getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void mergedCells() throws Exception {
        Path file = tempDir.resolve("merged.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Merged");
            sheet.cell(0, 0).set("Merged Header");
            sheet.merge(0, 0, 0, 4); // A1:E1

            sheet.cell(2, 1).set("Another Merge");
            sheet.merge(2, 1, 4, 3); // B3:D5
            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);

            assertEquals(2, sheet.getNumMergedRegions());

            var region1 = sheet.getMergedRegion(0);
            assertEquals(0, region1.getFirstRow());
            assertEquals(0, region1.getLastRow());
            assertEquals(0, region1.getFirstColumn());
            assertEquals(4, region1.getLastColumn());

            var region2 = sheet.getMergedRegion(1);
            assertEquals(2, region2.getFirstRow());
            assertEquals(4, region2.getLastRow());
            assertEquals(1, region2.getFirstColumn());
            assertEquals(3, region2.getLastColumn());

            assertEquals("Merged Header", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Another Merge", sheet.getRow(2).getCell(1).getStringCellValue());
        }
    }

    @Test
    void largeDataset() throws Exception {
        Path file = tempDir.resolve("large.xlsx");
        int rows = 1000;
        int cols = 10;

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Large");
            for (int r = 0; r < rows; r++) {
                for (int c = 0; c < cols; c++) {
                    sheet.cell(r, c).set("R" + r + "C" + c);
                }
            }
            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);

            // Spot check
            assertEquals("R0C0", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("R0C9", sheet.getRow(0).getCell(9).getStringCellValue());
            assertEquals("R500C5", sheet.getRow(500).getCell(5).getStringCellValue());
            assertEquals("R999C9", sheet.getRow(999).getCell(9).getStringCellValue());

            // Count rows
            int rowCount = 0;
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                if (sheet.getRow(i) != null) rowCount++;
            }
            assertEquals(rows, rowCount);
        }
    }

    @Test
    void styledCells() throws Exception {
        Path file = tempDir.resolve("styled.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Styled");

            int boldStyle = wb.addStyle(Style.builder().bold(true).build());
            int colorStyle = wb.addStyle(Style.builder().fill(0xFFFF0000).build());

            sheet.cell(0, 0).set("Bold").style(boldStyle);
            sheet.cell(0, 1).set("Colored").style(colorStyle);
            sheet.cell(0, 2).set("Normal");

            wb.save(file);
        }

        // Verify with POI - just check it opens and has data
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);

            assertEquals("Bold", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Colored", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals("Normal", sheet.getRow(0).getCell(2).getStringCellValue());

            // Verify bold style
            assertTrue(sheet.getRow(0).getCell(0).getCellStyle().getFont().getBold());
        }
    }

    @Test
    void protectedSheet() throws Exception {
        Path file = tempDir.resolve("protected.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Protected");
            sheet.cell(0, 0).set("Protected Data");
            sheet.cell(1, 0).set("Cannot edit");

            // Protect the sheet
            sheet.protectionManager().protect("password123".toCharArray(), SheetProtection.builder()
                .formatCells(false)
                .insertRows(false)
                .deleteRows(false)
                .sort(true)
                .autoFilter(true)
                .build());

            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);

            // Check sheet is protected
            assertNotNull(sheet.getCTWorksheet().getSheetProtection());

            // Data should still be readable
            assertEquals("Protected Data", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Cannot edit", sheet.getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    void protectedSheetWithPermissions() throws Exception {
        Path file = tempDir.resolve("protected-perms.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Permissions");
            sheet.cell(0, 0).set("Header");

            // Protect with specific permissions
            sheet.protectionManager().protect("secret".toCharArray(), SheetProtection.builder()
                .selectLockedCells(true)
                .selectUnlockedCells(true)
                .formatCells(true)
                .formatColumns(true)
                .formatRows(true)
                .insertRows(true)
                .deleteRows(true)
                .sort(true)
                .autoFilter(true)
                .pivotTables(true)
                .build());

            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = wb.getSheetAt(0);
            assertNotNull(sheet.getCTWorksheet().getSheetProtection());
            assertEquals("Header", sheet.getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void encryptedFile() throws Exception {
        Path file = tempDir.resolve("encrypted.xlsx");
        String password = "testPassword123";

        // Write with litexl (encrypted)
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Secret");
            sheet.cell(0, 0).set("Confidential Data");
            sheet.cell(1, 0).set(12345);
            sheet.cell(2, 0).set(true);

            wb.save(file, EncryptionOptions.aes256(password));
        }

        // Verify with POI - open encrypted file
        try (POIFSFileSystem fs = new POIFSFileSystem(file.toFile())) {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor decryptor = Decryptor.getInstance(info);

            assertTrue(decryptor.verifyPassword(password));

            try (InputStream is = decryptor.getDataStream(fs);
                 XSSFWorkbook wb = new XSSFWorkbook(is)) {

                XSSFSheet sheet = wb.getSheetAt(0);
                assertEquals("Secret", sheet.getSheetName());
                assertEquals("Confidential Data", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals(12345.0, sheet.getRow(1).getCell(0).getNumericCellValue());
                assertTrue(sheet.getRow(2).getCell(0).getBooleanCellValue());
            }
        }
    }

    @Test
    void encryptedFileAes128() throws Exception {
        Path file = tempDir.resolve("encrypted-aes128.xlsx");
        String password = "aes128pass";

        // Write with litexl (AES-128)
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("AES128");
            sheet.cell(0, 0).set("AES-128 Encrypted");

            wb.save(file, EncryptionOptions.aes128(password));
        }

        // Verify with POI
        try (POIFSFileSystem fs = new POIFSFileSystem(file.toFile())) {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor decryptor = Decryptor.getInstance(info);

            assertTrue(decryptor.verifyPassword(password));

            try (InputStream is = decryptor.getDataStream(fs);
                 XSSFWorkbook wb = new XSSFWorkbook(is)) {

                assertEquals("AES-128 Encrypted", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            }
        }
    }

    @Test
    void multipleSheetsWithData() throws Exception {
        Path file = tempDir.resolve("multi-data.xlsx");

        // Write with litexl
        try (Workbook wb = Workbook.create()) {
            // Sheet 1: Numbers
            Sheet numbers = wb.addSheet("Numbers");
            for (int i = 0; i < 10; i++) {
                numbers.cell(i, 0).set(i);
                numbers.cell(i, 1).set(i * i);
                numbers.cell(i, 2).setFormula("A" + (i + 1) + "*B" + (i + 1));
            }

            // Sheet 2: Strings
            Sheet strings = wb.addSheet("Strings");
            strings.cell(0, 0).set("Name");
            strings.cell(0, 1).set("City");
            strings.cell(1, 0).set("Alice");
            strings.cell(1, 1).set("Seoul");
            strings.cell(2, 0).set("Bob");
            strings.cell(2, 1).set("Tokyo");

            // Sheet 3: Mixed with styles
            Sheet mixed = wb.addSheet("Mixed");
            int boldStyle = wb.addStyle(Style.builder().bold(true).build());
            mixed.cell(0, 0).set("Header").style(boldStyle);
            mixed.cell(1, 0).set(100);
            mixed.cell(2, 0).set(true);

            // Sheet 4: Empty sheet
            wb.addSheet("Empty");

            // Sheet 5: Single cell
            Sheet single = wb.addSheet("Single");
            single.cell(0, 0).set("Only cell");

            wb.save(file);
        }

        // Verify with POI
        try (FileInputStream fis = new FileInputStream(file.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            assertEquals(5, wb.getNumberOfSheets());

            // Sheet 1: Numbers
            XSSFSheet numbers = wb.getSheet("Numbers");
            assertNotNull(numbers);
            assertEquals(0.0, numbers.getRow(0).getCell(0).getNumericCellValue());
            assertEquals(81.0, numbers.getRow(9).getCell(1).getNumericCellValue()); // 9*9
            assertEquals("A10*B10", numbers.getRow(9).getCell(2).getCellFormula());

            // Sheet 2: Strings
            XSSFSheet strings = wb.getSheet("Strings");
            assertNotNull(strings);
            assertEquals("Alice", strings.getRow(1).getCell(0).getStringCellValue());
            assertEquals("Tokyo", strings.getRow(2).getCell(1).getStringCellValue());

            // Sheet 3: Mixed
            XSSFSheet mixed = wb.getSheet("Mixed");
            assertNotNull(mixed);
            assertEquals("Header", mixed.getRow(0).getCell(0).getStringCellValue());
            assertTrue(mixed.getRow(0).getCell(0).getCellStyle().getFont().getBold());
            assertEquals(100.0, mixed.getRow(1).getCell(0).getNumericCellValue());
            assertTrue(mixed.getRow(2).getCell(0).getBooleanCellValue());

            // Sheet 4: Empty
            XSSFSheet empty = wb.getSheet("Empty");
            assertNotNull(empty);
            assertNull(empty.getRow(0));

            // Sheet 5: Single
            XSSFSheet single = wb.getSheet("Single");
            assertNotNull(single);
            assertEquals("Only cell", single.getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void encryptedMultipleSheets() throws Exception {
        Path file = tempDir.resolve("encrypted-multi.xlsx");
        String password = "multiPass";

        // Write encrypted workbook with multiple sheets
        try (Workbook wb = Workbook.create()) {
            Sheet sheet1 = wb.addSheet("Data1");
            sheet1.cell(0, 0).set("First Sheet Data");

            Sheet sheet2 = wb.addSheet("Data2");
            sheet2.cell(0, 0).set("Second Sheet Data");
            sheet2.cell(1, 0).set(999);

            wb.save(file, EncryptionOptions.aes256(password));
        }

        // Verify with POI
        try (POIFSFileSystem fs = new POIFSFileSystem(file.toFile())) {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor decryptor = Decryptor.getInstance(info);
            assertTrue(decryptor.verifyPassword(password));

            try (InputStream is = decryptor.getDataStream(fs);
                 XSSFWorkbook wb = new XSSFWorkbook(is)) {

                assertEquals(2, wb.getNumberOfSheets());
                assertEquals("First Sheet Data", wb.getSheet("Data1").getRow(0).getCell(0).getStringCellValue());
                assertEquals("Second Sheet Data", wb.getSheet("Data2").getRow(0).getCell(0).getStringCellValue());
                assertEquals(999.0, wb.getSheet("Data2").getRow(1).getCell(0).getNumericCellValue());
            }
        }
    }
}
