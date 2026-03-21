package com.beingidly.litexl;

import com.beingidly.litexl.crypto.WorkbookProtection;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

class WorkbookProtectionManagerTest {

    @TempDir
    Path tempDir;

    @Test
    void protectStructure_defaultOptions() {
        try (Workbook wb = Workbook.create()) {
            wb.protectStructure(WorkbookProtection.defaults());
            assertTrue(wb.isStructureProtected());
            assertNotNull(wb.structureProtection());
        }
    }

    @Test
    void protectStructure_withPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.protectStructure("password".toCharArray(), WorkbookProtection.defaults());
            assertTrue(wb.isStructureProtected());
        }
    }

    @Test
    void protectStructure_nullPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.workbookProtectionManager().protect(null, WorkbookProtection.defaults());
            assertTrue(wb.isStructureProtected());
            assertNull(wb.workbookProtectionManager().passwordInfo());
        }
    }

    @Test
    void protectStructure_emptyPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.workbookProtectionManager().protect(new char[0], WorkbookProtection.defaults());
            assertTrue(wb.isStructureProtected());
            assertNull(wb.workbookProtectionManager().passwordInfo());
        }
    }

    @Test
    void unprotectStructure_noPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.protectStructure(WorkbookProtection.defaults());
            assertTrue(wb.isStructureProtected());

            boolean result = wb.unprotectStructure();
            assertTrue(result);
            assertFalse(wb.isStructureProtected());
        }
    }

    @Test
    void unprotectStructure_withCorrectPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.protectStructure("password".toCharArray(), WorkbookProtection.defaults());
            assertTrue(wb.isStructureProtected());

            boolean result = wb.unprotectStructure("password".toCharArray());
            assertTrue(result);
            assertFalse(wb.isStructureProtected());
        }
    }

    @Test
    void unprotectStructure_withWrongPassword() {
        try (Workbook wb = Workbook.create()) {
            wb.protectStructure("password".toCharArray(), WorkbookProtection.defaults());
            assertTrue(wb.isStructureProtected());

            boolean result = wb.unprotectStructure("wrong".toCharArray());
            assertFalse(result);
            assertTrue(wb.isStructureProtected());
        }
    }

    @Test
    void unprotectStructure_withoutPassword_whenPasswordSet() {
        try (Workbook wb = Workbook.create()) {
            wb.protectStructure("password".toCharArray(), WorkbookProtection.defaults());
            assertTrue(wb.isStructureProtected());

            boolean result = wb.unprotectStructure();
            assertFalse(result);
            assertTrue(wb.isStructureProtected());
        }
    }

    @Test
    void unprotectStructure_whenNotProtected() {
        try (Workbook wb = Workbook.create()) {
            assertFalse(wb.isStructureProtected());

            boolean result = wb.unprotectStructure("password".toCharArray());
            assertTrue(result);
        }
    }

    @Test
    void protectStructure_customOptions() {
        try (Workbook wb = Workbook.create()) {
            WorkbookProtection options = WorkbookProtection.builder()
                    .lockStructure(false)
                    .lockWindows(true)
                    .build();

            wb.protectStructure(options);
            assertTrue(wb.isStructureProtected());

            WorkbookProtection prot = wb.structureProtection();
            assertNotNull(prot);
            assertFalse(prot.lockStructure());
            assertTrue(prot.lockWindows());
        }
    }

    @Test
    void structureProtection_returnsNullWhenNotProtected() {
        try (Workbook wb = Workbook.create()) {
            assertNull(wb.structureProtection());
        }
    }

    @Test
    void saveProtectedWorkbook_structure() {
        Path file = tempDir.resolve("protected_structure.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.cell(0, 0).set("Data");
            wb.protectStructure(WorkbookProtection.defaults());
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertTrue(wb.isStructureProtected());
            WorkbookProtection prot = wb.structureProtection();
            assertNotNull(prot);
            assertTrue(prot.lockStructure());
            assertFalse(prot.lockWindows());

            Sheet sheet = wb.getSheet(0);
            assertEquals("Data", sheet.cell(0, 0).string());
        }
    }

    @Test
    void saveProtectedWorkbook_structureWithPassword() {
        Path file = tempDir.resolve("protected_structure_password.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.cell(0, 0).set("Data");
            wb.protectStructure("secret".toCharArray(), WorkbookProtection.defaults());
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertTrue(wb.isStructureProtected());
            assertNotNull(wb.workbookProtectionManager().passwordInfo());

            Sheet sheet = wb.getSheet(0);
            assertEquals("Data", sheet.cell(0, 0).string());
        }
    }

    @Test
    void saveProtectedWorkbook_customOptions() {
        Path file = tempDir.resolve("protected_custom.xlsx");

        WorkbookProtection options = WorkbookProtection.builder()
                .lockStructure(true)
                .lockWindows(true)
                .build();

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.cell(0, 0).set("Data");
            wb.protectStructure("pass".toCharArray(), options);
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            assertTrue(wb.isStructureProtected());
            WorkbookProtection prot = wb.structureProtection();
            assertNotNull(prot);
            assertTrue(prot.lockStructure());
            assertTrue(prot.lockWindows());
        }
    }

    @Test
    void saveBothProtections() {
        Path file = tempDir.resolve("both_protections.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Test");
            sheet.cell(0, 0).set("Data");
            wb.setWriteProtection("wp_pass".toCharArray(), "Admin");
            wb.protectStructure("struct_pass".toCharArray(), WorkbookProtection.defaults());
            wb.save(file);
        }

        try (Workbook wb = Workbook.open(file)) {
            // Write protection
            assertTrue(wb.isWriteProtected());
            assertEquals("Admin", wb.writeProtection().userName());
            assertTrue(wb.writeProtection().readOnlyRecommended());
            assertNotNull(wb.writeProtectionManager().passwordInfo());

            // Structure protection
            assertTrue(wb.isStructureProtected());
            assertTrue(wb.structureProtection().lockStructure());
            assertNotNull(wb.workbookProtectionManager().passwordInfo());

            Sheet sheet = wb.getSheet(0);
            assertEquals("Data", sheet.cell(0, 0).string());
        }
    }
}
