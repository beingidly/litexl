package com.beingidly.litexl;

import com.beingidly.litexl.crypto.EncryptionOptions;
import com.beingidly.litexl.style.Style;
import com.beingidly.litexl.style.BorderStyle;
import com.beingidly.litexl.style.HAlign;
import com.beingidly.litexl.style.VAlign;
import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.*;

class WorkbookTest {

    @TempDir
    Path tempDir;

    @Test
    void createEmptyWorkbook() {
        try (Workbook wb = Workbook.create()) {
            assertEquals(0, wb.sheetCount());
            assertNotNull(wb.sheets());
            assertTrue(wb.sheets().isEmpty());
        }
    }

    @Test
    void addSheet() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");

            assertEquals(1, wb.sheetCount());
            assertEquals("Test", sheet.name());
            assertEquals(0, sheet.index());
            assertSame(sheet, wb.getSheet(0));
            assertSame(sheet, wb.getSheet("Test"));
        }
    }

    @Test
    void addSheetDuplicateName() {
        try (Workbook wb = Workbook.create()) {
            wb.addSheet("Test");
            assertThrows(IllegalArgumentException.class, () -> wb.addSheet("Test"));
            assertThrows(IllegalArgumentException.class, () -> wb.addSheet("TEST"));
        }
    }

    @Test
    void removeSheet() {
        try (Workbook wb = Workbook.create()) {
            wb.addSheet("Sheet1");
            wb.addSheet("Sheet2");
            assertEquals(2, wb.sheetCount());

            wb.removeSheet(0);
            assertEquals(1, wb.sheetCount());
            assertEquals("Sheet2", wb.getSheet(0).name());
        }
    }

    @Test
    void setCellValues() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");

            sheet.cell(0, 0).set("Hello");
            sheet.cell(0, 1).set(42.5);
            sheet.cell(0, 2).set(true);
            sheet.cell(0, 3).set(LocalDateTime.of(2024, 1, 15, 10, 30));
            sheet.cell(0, 4).setFormula("SUM(B1:B10)");

            assertEquals(CellType.STRING, sheet.cell(0, 0).type());
            assertEquals("Hello", sheet.cell(0, 0).string());

            assertEquals(CellType.NUMBER, sheet.cell(0, 1).type());
            assertEquals(42.5, sheet.cell(0, 1).number());

            assertEquals(CellType.BOOLEAN, sheet.cell(0, 2).type());
            assertTrue(sheet.cell(0, 2).bool());

            assertEquals(CellType.DATE, sheet.cell(0, 3).type());
            assertEquals(LocalDateTime.of(2024, 1, 15, 10, 30), sheet.cell(0, 3).date());

            assertEquals(CellType.FORMULA, sheet.cell(0, 4).type());
            assertEquals("SUM(B1:B10)", sheet.cell(0, 4).formula());
        }
    }

    @Test
    void cellChaining() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");

            Style style = Style.builder()
                .font("Arial", 12)
                .bold(true)
                .fill(0xFFCCCCCC)
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .border(BorderStyle.THIN, 0xFF000000)
                .build();

            int styleId = wb.addStyle(style);

            sheet.cell(0, 0)
                .set("Header")
                .style(styleId);

            assertEquals("Header", sheet.cell(0, 0).string());
            assertEquals(styleId, sheet.cell(0, 0).styleId());
        }
    }

    @Test
    void saveAndOpen() throws Exception {
        Path file = tempDir.resolve("test.xlsx");

        // Create and save
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Hello");
            sheet.cell(0, 1).set(123.45);
            sheet.cell(1, 0).set(true);
            wb.save(file);
        }

        // Open and verify
        try (Workbook wb = Workbook.open(file)) {
            assertEquals(1, wb.sheetCount());

            Sheet sheet = wb.getSheet(0);
            assertEquals("Data", sheet.name());

            assertEquals("Hello", sheet.getCell(0, 0).string());
            assertEquals(123.45, sheet.getCell(0, 1).number(), 0.001);
            assertTrue(sheet.getCell(1, 0).bool());
        }
    }

    @Test
    void mergeCells() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");

            sheet.cell(0, 0).set("Merged");
            sheet.merge(0, 0, 0, 2);

            assertEquals(1, sheet.mergedCells().size());
            assertEquals("A1:C1", sheet.mergedCells().get(0).toRange().toRef());
        }
    }

    @Test
    void autoFilter() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");

            sheet.setAutoFilter(0, 0, 10, 5);

            assertNotNull(sheet.autoFilter());
            assertEquals("A1:F11", sheet.autoFilter().range().toRef());

            sheet.clearAutoFilter();
            assertNull(sheet.autoFilter());
        }
    }

    @Test
    void closedWorkbookThrows() {
        Workbook wb = Workbook.create();
        wb.close();

        assertThrows(IllegalStateException.class, () -> wb.addSheet("Test"));
        assertThrows(IllegalStateException.class, () -> wb.getSheet(0));
    }

    @Test
    void addSheet_emptyNameThrows() {
        try (Workbook wb = Workbook.create()) {
            assertThrows(IllegalArgumentException.class, () -> wb.addSheet(""));
        }
    }

    @Test
    void getSheet_byIndex_outOfRange() {
        try (Workbook wb = Workbook.create()) {
            wb.addSheet("Sheet1");
            assertNull(wb.getSheet(-1));
            assertNull(wb.getSheet(1));
        }
    }

    @Test
    void getSheet_byName_notFound() {
        try (Workbook wb = Workbook.create()) {
            wb.addSheet("Sheet1");
            assertNull(wb.getSheet("NonExistent"));
        }
    }

    @Test
    void getSheet_byName_caseInsensitive() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("MySheet");
            assertSame(sheet, wb.getSheet("mysheet"));
            assertSame(sheet, wb.getSheet("MYSHEET"));
        }
    }

    @Test
    void removeSheet_outOfRange() {
        try (Workbook wb = Workbook.create()) {
            wb.addSheet("Sheet1");
            assertThrows(IndexOutOfBoundsException.class, () -> wb.removeSheet(-1));
            assertThrows(IndexOutOfBoundsException.class, () -> wb.removeSheet(1));
        }
    }

    @Test
    void sharedStrings() {
        try (Workbook wb = Workbook.create()) {
            int idx1 = wb.addSharedString("Hello");
            int idx2 = wb.addSharedString("World");
            int idx3 = wb.addSharedString("Hello"); // duplicate

            assertEquals(0, idx1);
            assertEquals(1, idx2);
            assertEquals(0, idx3); // same as idx1

            assertEquals("Hello", wb.getSharedString(0));
            assertEquals("World", wb.getSharedString(1));
            assertNull(wb.getSharedString(-1));
            assertNull(wb.getSharedString(100));

            assertEquals(2, wb.sharedStrings().size());
        }
    }

    @Test
    void styles() {
        try (Workbook wb = Workbook.create()) {
            // Default style at index 0
            assertNotNull(wb.getStyle(0));
            assertNull(wb.getStyle(-1));
            assertNull(wb.getStyle(100));
            assertTrue(wb.styles().size() >= 1);
        }
    }

    @Test
    void toString_format() {
        try (Workbook wb = Workbook.create()) {
            assertEquals("Workbook[sheets=0]", wb.toString());
            wb.addSheet("Test");
            assertEquals("Workbook[sheets=1]", wb.toString());
        }
    }

    @Test
    void openNonExistentFile() {
        assertThrows(LitexlException.class,
            () -> Workbook.open(Path.of("/nonexistent/path/file.xlsx")));
    }

    @Test
    void closedWorkbook_getSheetByName() {
        Workbook wb = Workbook.create();
        wb.addSheet("Test");
        wb.close();
        assertThrows(IllegalStateException.class, () -> wb.getSheet("Test"));
    }

    @Test
    void closedWorkbook_removeSheet() {
        Workbook wb = Workbook.create();
        wb.addSheet("Test");
        wb.close();
        assertThrows(IllegalStateException.class, () -> wb.removeSheet(0));
    }

    @Test
    void closedWorkbook_addStyle() {
        Workbook wb = Workbook.create();
        wb.close();
        assertThrows(IllegalStateException.class, () -> wb.addStyle(Style.DEFAULT));
    }

    @Test
    void closedWorkbook_save() {
        Workbook wb = Workbook.create();
        wb.close();
        assertThrows(IllegalStateException.class, () -> wb.save(tempDir.resolve("test.xlsx")));
    }

    @Test
    void saveToOutputStream() throws Exception {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        // Create and save to stream
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("StreamData");
            sheet.cell(0, 0).set("Hello Stream");
            sheet.cell(0, 1).set(99.99);
            sheet.cell(1, 0).set(false);
            wb.save(baos);
        }

        // Verify we got data
        byte[] data = baos.toByteArray();
        assertTrue(data.length > 0, "Output stream should contain data");

        // Save to temp file and open to verify content
        Path tempFile = tempDir.resolve("stream-test.xlsx");
        Files.write(tempFile, data);

        try (Workbook wb = Workbook.open(tempFile)) {
            assertEquals(1, wb.sheetCount());
            Sheet sheet = wb.getSheet(0);
            assertEquals("StreamData", sheet.name());
            assertEquals("Hello Stream", sheet.getCell(0, 0).string());
            assertEquals(99.99, sheet.getCell(0, 1).number(), 0.001);
            assertFalse(sheet.getCell(1, 0).bool());
        }
    }

    @Test
    void saveToOutputStream_multipleSheets() throws Exception {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (Workbook wb = Workbook.create()) {
            Sheet sheet1 = wb.addSheet("Sheet1");
            sheet1.cell(0, 0).set("First");

            Sheet sheet2 = wb.addSheet("Sheet2");
            sheet2.cell(0, 0).set("Second");

            wb.save(baos);
        }

        Path tempFile = tempDir.resolve("multi-sheet.xlsx");
        Files.write(tempFile, baos.toByteArray());

        try (Workbook wb = Workbook.open(tempFile)) {
            assertEquals(2, wb.sheetCount());
            assertEquals("First", wb.getSheet("Sheet1").getCell(0, 0).string());
            assertEquals("Second", wb.getSheet("Sheet2").getCell(0, 0).string());
        }
    }

    @Test
    void saveToOutputStream_closedWorkbook() {
        Workbook wb = Workbook.create();
        wb.close();
        assertThrows(IllegalStateException.class, () -> wb.save(new ByteArrayOutputStream()));
    }

    @Test
    void saveToOutputStream_withEncryption() throws Exception {
        String password = "testPass123";
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        // Create and save encrypted to stream
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("EncryptedStream");
            sheet.cell(0, 0).set("Secret Data");
            sheet.cell(0, 1).set(12345.67);
            wb.save(baos, EncryptionOptions.aes256(password));
        }

        // Verify we got data
        byte[] data = baos.toByteArray();
        assertTrue(data.length > 0, "Output stream should contain encrypted data");

        // Save to temp file and open with password to verify
        Path tempFile = tempDir.resolve("encrypted-stream.xlsx");
        Files.write(tempFile, data);

        try (Workbook wb = Workbook.open(tempFile, password)) {
            assertEquals(1, wb.sheetCount());
            Sheet sheet = wb.getSheet(0);
            assertEquals("EncryptedStream", sheet.name());
            assertEquals("Secret Data", sheet.getCell(0, 0).string());
            assertEquals(12345.67, sheet.getCell(0, 1).number(), 0.001);
        }
    }
}
