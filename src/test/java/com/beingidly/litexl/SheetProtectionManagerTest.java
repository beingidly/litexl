package com.beingidly.litexl;

import com.beingidly.litexl.crypto.SheetProtection;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

class SheetProtectionManagerTest {

    @TempDir
    Path tempDir;

    @Test
    void protect_defaultOptions() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.protect(SheetProtection.defaults());
            assertTrue(sheet.isProtected());
            assertNotNull(sheet.protection());
        }
    }

    @Test
    void protect_withPassword() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.protectionManager().protect("password".toCharArray(), SheetProtection.defaults());
            assertTrue(sheet.isProtected());
        }
    }

    @Test
    void protect_nullPassword() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.protectionManager().protect(null, SheetProtection.defaults());
            assertTrue(sheet.isProtected());
        }
    }

    @Test
    void protect_emptyPassword() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.protectionManager().protect(new char[0], SheetProtection.defaults());
            assertTrue(sheet.isProtected());
        }
    }

    @Test
    void unprotect_noPassword() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.protect(SheetProtection.defaults());
            assertTrue(sheet.isProtected());

            boolean result = sheet.protectionManager().unprotect();
            assertTrue(result);
            assertFalse(sheet.isProtected());
        }
    }

    @Test
    void unprotect_withCorrectPassword() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.protectionManager().protect("password".toCharArray(), SheetProtection.defaults());
            assertTrue(sheet.isProtected());

            boolean result = sheet.protectionManager().unprotect("password".toCharArray());
            assertTrue(result);
            assertFalse(sheet.isProtected());
        }
    }

    @Test
    void unprotect_withWrongPassword() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.protectionManager().protect("password".toCharArray(), SheetProtection.defaults());
            assertTrue(sheet.isProtected());

            boolean result = sheet.protectionManager().unprotect("wrong".toCharArray());
            assertFalse(result);
            assertTrue(sheet.isProtected());
        }
    }

    @Test
    void unprotect_withoutPassword_whenPasswordSet() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.protectionManager().protect("password".toCharArray(), SheetProtection.defaults());
            assertTrue(sheet.isProtected());

            boolean result = sheet.protectionManager().unprotect();
            assertFalse(result);
            assertTrue(sheet.isProtected());
        }
    }

    @Test
    void unprotect_whenNotProtected() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            assertFalse(sheet.isProtected());

            boolean result = sheet.protectionManager().unprotect("password".toCharArray());
            assertTrue(result);
        }
    }

    @Test
    void protect_customOptions() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            SheetProtection options = SheetProtection.builder()
                    .formatCells(true)
                    .insertRows(true)
                    .deleteRows(true)
                    .sort(true)
                    .build();

            sheet.protect(options);
            assertTrue(sheet.isProtected());

            SheetProtection protection = sheet.protection();
            assertNotNull(protection);
            assertTrue(protection.formatCells());
            assertTrue(protection.insertRows());
            assertTrue(protection.deleteRows());
            assertTrue(protection.sort());
            assertFalse(protection.formatColumns());
            assertFalse(protection.autoFilter());
        }
    }

    @Test
    void protection_returnsNullWhenNotProtected() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            assertNull(sheet.protection());
        }
    }

    @Test
    void saveProtectedSheet_noPassword() {
        // Verifies that protected sheet can be saved without errors
        // Note: XlsxReader doesn't currently parse sheetProtection element back
        Path file = tempDir.resolve("protected_no_password.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Protected");
            sheet.cell(0, 0).set("Data");
            sheet.protect(SheetProtection.defaults());
            wb.save(file);
        }

        // File should be readable
        try (Workbook wb = Workbook.open(file)) {
            Sheet sheet = wb.getSheet(0);
            assertEquals("Data", sheet.cell(0, 0).string());
        }
    }

    @Test
    void saveProtectedSheet_withPassword() {
        // Verifies that password-protected sheet can be saved without errors
        Path file = tempDir.resolve("protected_with_password.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Protected");
            sheet.cell(0, 0).set("Data");
            sheet.protectionManager().protect("secret".toCharArray(), SheetProtection.defaults());
            wb.save(file);
        }

        // File should be readable
        try (Workbook wb = Workbook.open(file)) {
            Sheet sheet = wb.getSheet(0);
            assertEquals("Data", sheet.cell(0, 0).string());
        }
    }

    @Test
    void saveProtectedSheet_withCustomOptions() {
        // Verifies that sheet with custom protection options can be saved
        Path file = tempDir.resolve("protected_custom.xlsx");

        SheetProtection options = SheetProtection.builder()
                .selectLockedCells(false)
                .formatCells(true)
                .insertColumns(true)
                .autoFilter(true)
                .build();

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Protected");
            sheet.cell(0, 0).set("Data");
            sheet.protect(options);
            wb.save(file);
        }

        // File should be readable
        try (Workbook wb = Workbook.open(file)) {
            Sheet sheet = wb.getSheet(0);
            assertEquals("Data", sheet.cell(0, 0).string());
        }
    }

}
