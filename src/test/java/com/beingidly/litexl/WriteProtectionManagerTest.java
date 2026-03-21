package com.beingidly.litexl;

import com.beingidly.litexl.crypto.WriteProtection;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

class WriteProtectionManagerTest {

    @TempDir
    Path tempDir;

    @Test
    void protect_withPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.setWriteProtection("password".toCharArray(), "Admin");
            assertTrue(wb.isWriteProtected());

            WriteProtection wp = wb.writeProtection();
            assertNotNull(wp);
            assertTrue(wp.readOnlyRecommended());
            assertEquals("Admin", wp.userName());
        }
    }

    @Test
    void protect_withoutPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.setWriteProtection("User1");
            assertTrue(wb.isWriteProtected());

            WriteProtection wp = wb.writeProtection();
            assertNotNull(wp);
            assertTrue(wp.readOnlyRecommended());
            assertEquals("User1", wp.userName());
        }
    }

    @Test
    void protect_nullPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.writeProtectionManager().protect(null, "User1");
            assertTrue(wb.isWriteProtected());
            assertNull(wb.writeProtectionManager().passwordInfo());
        }
    }

    @Test
    void protect_emptyPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.writeProtectionManager().protect(new char[0], "User1");
            assertTrue(wb.isWriteProtected());
            assertNull(wb.writeProtectionManager().passwordInfo());
        }
    }

    @Test
    void removeWriteProtection() {
        try (Workbook wb = Workbook.create()) {
            wb.setWriteProtection("password".toCharArray(), "Admin");
            assertTrue(wb.isWriteProtected());

            wb.removeWriteProtection();
            assertFalse(wb.isWriteProtected());
            assertNull(wb.writeProtection());
        }
    }

    @Test
    void isWriteProtected_defaultFalse() {
        try (Workbook wb = Workbook.create()) {
            assertFalse(wb.isWriteProtected());
        }
    }

    @Test
    void writeProtection_returnsNullWhenNotSet() {
        try (Workbook wb = Workbook.create()) {
            assertNull(wb.writeProtection());
        }
    }

    @Test
    void saveWriteProtected_withPassword() {
        Path file = tempDir.resolve("write_protected_password.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.cell(0, 0).set("Data");
            wb.setWriteProtection("secret".toCharArray(), "Admin");
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertTrue(wb.isWriteProtected());
            WriteProtection wp = wb.writeProtection();
            assertNotNull(wp);
            assertTrue(wp.readOnlyRecommended());
            assertEquals("Admin", wp.userName());
            assertNotNull(wb.writeProtectionManager().passwordInfo());

            Sheet sheet = wb.getSheet(0);
            assertEquals("Data", sheet.cell(0, 0).string());
        }
    }

    @Test
    void saveWriteProtected_withoutPassword() {
        Path file = tempDir.resolve("write_protected_no_password.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.cell(0, 0).set("Data");
            wb.setWriteProtection("User1");
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertTrue(wb.isWriteProtected());
            WriteProtection wp = wb.writeProtection();
            assertNotNull(wp);
            assertTrue(wp.readOnlyRecommended());
            assertEquals("User1", wp.userName());
            assertNull(wb.writeProtectionManager().passwordInfo());
        }
    }

    @Test
    void saveWriteProtected_roundtripPreservesData() {
        Path file = tempDir.resolve("write_protected_roundtrip.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            sheet.cell(0, 0).set("Hello");
            sheet.cell(1, 0).set(42.0);
            wb.setWriteProtection("pass123".toCharArray(), "TestUser");
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertTrue(wb.isWriteProtected());
            assertEquals("TestUser", wb.writeProtection().userName());

            Sheet sheet = wb.getSheet(0);
            assertEquals("Hello", sheet.cell(0, 0).string());
            assertEquals(42.0, sheet.cell(1, 0).number());
        }
    }
}
