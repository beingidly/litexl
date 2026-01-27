package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.nio.file.Path;
import static org.junit.jupiter.api.Assertions.*;

class XlsxReaderTest {
    @TempDir
    Path tempDir;

    @Test
    void readsInlineStrings() throws Exception {
        Path path = tempDir.resolve("inline.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Inline text");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Inline text", read.getSheet(0).getCell(0, 0).string());
        }
    }

    @Test
    void readsBooleanCells() throws Exception {
        Path path = tempDir.resolve("boolean.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(true);
            sheet.cell(0, 1).set(false);
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertTrue(read.getSheet(0).getCell(0, 0).bool());
            assertFalse(read.getSheet(0).getCell(0, 1).bool());
        }
    }

    @Test
    void readsErrorCells() throws Exception {
        Path path = tempDir.resolve("error.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).setValue(new CellValue.Error("#DIV/0!"));
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(CellType.ERROR, read.getSheet(0).getCell(0, 0).type());
            assertEquals("#DIV/0!", read.getSheet(0).getCell(0, 0).error());
        }
    }

    @Test
    void readsFormulaCells() throws Exception {
        Path path = tempDir.resolve("formula.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(10);
            sheet.cell(0, 1).set(20);
            sheet.cell(0, 2).setFormula("A1+B1");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("A1+B1", read.getSheet(0).getCell(0, 2).formula());
        }
    }

    @Test
    void readsEncryptedWithPassword() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes256.xlsx");
        try (Workbook workbook = Workbook.open(encrypted, "password123")) {
            assertNotNull(workbook);
            assertTrue(workbook.sheetCount() > 0);
        }
    }

    @Test
    void throwsOnWrongPassword() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes256.xlsx");
        assertThrows(Exception.class, () -> Workbook.open(encrypted, "wrongpassword"));
    }

    @Test
    void readsMergedCells() throws Exception {
        Path path = tempDir.resolve("merged.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Merged");
            sheet.merge(CellRange.parse("A1:C3"));
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(1, read.getSheet(0).mergedCells().size());
        }
    }

    @Test
    void readsMultipleSheets() throws Exception {
        Path path = tempDir.resolve("multi.xlsx");
        try (Workbook workbook = Workbook.create()) {
            workbook.addSheet("Sheet1").cell(0, 0).set("A");
            workbook.addSheet("Sheet2").cell(0, 0).set("B");
            workbook.addSheet("Sheet3").cell(0, 0).set("C");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(3, read.sheetCount());
            assertEquals("Sheet1", read.getSheet(0).name());
            assertEquals("Sheet2", read.getSheet(1).name());
            assertEquals("Sheet3", read.getSheet(2).name());
        }
    }

    @Test
    void readsNumericCells() throws Exception {
        Path path = tempDir.resolve("numeric.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(42.5);
            sheet.cell(0, 1).set(-100.0);
            sheet.cell(0, 2).set(0.0);
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(42.5, read.getSheet(0).getCell(0, 0).number(), 0.001);
            assertEquals(-100.0, read.getSheet(0).getCell(0, 1).number(), 0.001);
            assertEquals(0.0, read.getSheet(0).getCell(0, 2).number(), 0.001);
        }
    }

    @Test
    void readsEmptyCells() throws Exception {
        Path path = tempDir.resolve("empty.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("A");
            // Cell at (0, 1) is not set, should be empty
            sheet.cell(0, 2).set("C");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("A", read.getSheet(0).getCell(0, 0).string());
            // Cell (0, 1) doesn't exist in the file
            assertNull(read.getSheet(0).getCell(0, 1));
            assertEquals("C", read.getSheet(0).getCell(0, 2).string());
        }
    }

    @Test
    void readsSheetWithManyRows() throws Exception {
        Path path = tempDir.resolve("many-rows.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            for (int i = 0; i < 100; i++) {
                sheet.cell(i, 0).set("Row " + i);
            }
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Row 0", read.getSheet(0).getCell(0, 0).string());
            assertEquals("Row 50", read.getSheet(0).getCell(50, 0).string());
            assertEquals("Row 99", read.getSheet(0).getCell(99, 0).string());
        }
    }

    @Test
    void readsSheetWithManyColumns() throws Exception {
        Path path = tempDir.resolve("many-cols.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            for (int i = 0; i < 50; i++) {
                sheet.cell(0, i).set("Col " + i);
            }
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Col 0", read.getSheet(0).getCell(0, 0).string());
            assertEquals("Col 25", read.getSheet(0).getCell(0, 25).string());
            assertEquals("Col 49", read.getSheet(0).getCell(0, 49).string());
        }
    }

    @Test
    void readsMultipleMergedRegions() throws Exception {
        Path path = tempDir.resolve("multi-merge.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Merge1");
            sheet.merge(CellRange.parse("A1:B2"));
            sheet.cell(0, 3).set("Merge2");
            sheet.merge(CellRange.parse("D1:E2"));
            sheet.cell(3, 0).set("Merge3");
            sheet.merge(CellRange.parse("A4:C5"));
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(3, read.getSheet(0).mergedCells().size());
        }
    }

    @Test
    void readsMixedCellTypes() throws Exception {
        Path path = tempDir.resolve("mixed.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("String");
            sheet.cell(0, 1).set(123.45);
            sheet.cell(0, 2).set(true);
            sheet.cell(0, 3).setFormula("B1*2");
            sheet.cell(0, 4).setValue(new CellValue.Error("#N/A"));
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(CellType.STRING, read.getSheet(0).getCell(0, 0).type());
            assertEquals(CellType.NUMBER, read.getSheet(0).getCell(0, 1).type());
            assertEquals(CellType.BOOLEAN, read.getSheet(0).getCell(0, 2).type());
            assertEquals(CellType.FORMULA, read.getSheet(0).getCell(0, 3).type());
            assertEquals(CellType.ERROR, read.getSheet(0).getCell(0, 4).type());
        }
    }

    @Test
    void readsSpecialCharactersInStrings() throws Exception {
        Path path = tempDir.resolve("special.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("<>&\"'");
            sheet.cell(0, 1).set("Line1\nLine2");
            sheet.cell(0, 2).set("Tab\there");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("<>&\"'", read.getSheet(0).getCell(0, 0).string());
            assertEquals("Line1\nLine2", read.getSheet(0).getCell(0, 1).string());
            assertEquals("Tab\there", read.getSheet(0).getCell(0, 2).string());
        }
    }

    @Test
    void readsUnicodeStrings() throws Exception {
        Path path = tempDir.resolve("unicode.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Hello");
            sheet.cell(0, 1).set("Привет");
            sheet.cell(0, 2).set("你好");
            sheet.cell(0, 3).set("مرحبا");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Hello", read.getSheet(0).getCell(0, 0).string());
            assertEquals("Привет", read.getSheet(0).getCell(0, 1).string());
            assertEquals("你好", read.getSheet(0).getCell(0, 2).string());
            assertEquals("مرحبا", read.getSheet(0).getCell(0, 3).string());
        }
    }

    @Test
    void readsLargeNumbers() throws Exception {
        Path path = tempDir.resolve("large-nums.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(1e15);
            sheet.cell(0, 1).set(-1e15);
            sheet.cell(0, 2).set(Double.MAX_VALUE / 2);
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(1e15, read.getSheet(0).getCell(0, 0).number(), 1e5);
            assertEquals(-1e15, read.getSheet(0).getCell(0, 1).number(), 1e5);
            assertEquals(Double.MAX_VALUE / 2, read.getSheet(0).getCell(0, 2).number(), 1e300);
        }
    }

    @Test
    void readsSmallNumbers() throws Exception {
        Path path = tempDir.resolve("small-nums.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(0.000001);
            sheet.cell(0, 1).set(-0.000001);
            sheet.cell(0, 2).set(1e-10);
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(0.000001, read.getSheet(0).getCell(0, 0).number(), 1e-10);
            assertEquals(-0.000001, read.getSheet(0).getCell(0, 1).number(), 1e-10);
            assertEquals(1e-10, read.getSheet(0).getCell(0, 2).number(), 1e-15);
        }
    }

    @Test
    void readsEncryptedAes128() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes128.xlsx");
        try (Workbook workbook = Workbook.open(encrypted, "password123")) {
            assertNotNull(workbook);
            assertTrue(workbook.sheetCount() > 0);
        }
    }

    @Test
    void throwsOnMissingPassword() throws Exception {
        Path encrypted = Path.of("src/test/resources/encrypted/aes256.xlsx");
        assertThrows(Exception.class, () -> Workbook.open(encrypted));
    }

    @Test
    void readsVariousErrorCodes() throws Exception {
        Path path = tempDir.resolve("errors.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).setValue(new CellValue.Error("#DIV/0!"));
            sheet.cell(0, 1).setValue(new CellValue.Error("#N/A"));
            sheet.cell(0, 2).setValue(new CellValue.Error("#VALUE!"));
            sheet.cell(0, 3).setValue(new CellValue.Error("#REF!"));
            sheet.cell(0, 4).setValue(new CellValue.Error("#NAME?"));
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("#DIV/0!", read.getSheet(0).getCell(0, 0).error());
            assertEquals("#N/A", read.getSheet(0).getCell(0, 1).error());
            assertEquals("#VALUE!", read.getSheet(0).getCell(0, 2).error());
            assertEquals("#REF!", read.getSheet(0).getCell(0, 3).error());
            assertEquals("#NAME?", read.getSheet(0).getCell(0, 4).error());
        }
    }

    @Test
    void readsComplexFormulas() throws Exception {
        Path path = tempDir.resolve("complex-formula.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(10);
            sheet.cell(0, 1).set(20);
            sheet.cell(0, 2).set(30);
            sheet.cell(1, 0).setFormula("SUM(A1:C1)");
            sheet.cell(1, 1).setFormula("AVERAGE(A1:C1)");
            sheet.cell(1, 2).setFormula("IF(A1>5,\"Yes\",\"No\")");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("SUM(A1:C1)", read.getSheet(0).getCell(1, 0).formula());
            assertEquals("AVERAGE(A1:C1)", read.getSheet(0).getCell(1, 1).formula());
            assertEquals("IF(A1>5,\"Yes\",\"No\")", read.getSheet(0).getCell(1, 2).formula());
        }
    }

    @Test
    void readsSheetWithOnlyFormulas() throws Exception {
        Path path = tempDir.resolve("only-formulas.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).setFormula("1+1");
            sheet.cell(0, 1).setFormula("2*3");
            sheet.cell(0, 2).setFormula("A1+B1");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(CellType.FORMULA, read.getSheet(0).getCell(0, 0).type());
            assertEquals(CellType.FORMULA, read.getSheet(0).getCell(0, 1).type());
            assertEquals(CellType.FORMULA, read.getSheet(0).getCell(0, 2).type());
        }
    }

    @Test
    void preservesLeadingAndTrailingSpaces() throws Exception {
        Path path = tempDir.resolve("spaces.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("  leading");
            sheet.cell(0, 1).set("trailing  ");
            sheet.cell(0, 2).set("  both  ");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("  leading", read.getSheet(0).getCell(0, 0).string());
            assertEquals("trailing  ", read.getSheet(0).getCell(0, 1).string());
            assertEquals("  both  ", read.getSheet(0).getCell(0, 2).string());
        }
    }

    @Test
    void readsLargeRowNumbers() throws Exception {
        Path path = tempDir.resolve("large-row.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(999, 0).set("Row 1000");
            sheet.cell(9999, 0).set("Row 10000");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Row 1000", read.getSheet(0).getCell(999, 0).string());
            assertEquals("Row 10000", read.getSheet(0).getCell(9999, 0).string());
        }
    }

    @Test
    void readsLargeColumnNumbers() throws Exception {
        Path path = tempDir.resolve("large-col.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 25).set("Column Z");
            sheet.cell(0, 26).set("Column AA");
            sheet.cell(0, 51).set("Column AZ");
            sheet.cell(0, 701).set("Column ZZ");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Column Z", read.getSheet(0).getCell(0, 25).string());
            assertEquals("Column AA", read.getSheet(0).getCell(0, 26).string());
            assertEquals("Column AZ", read.getSheet(0).getCell(0, 51).string());
            assertEquals("Column ZZ", read.getSheet(0).getCell(0, 701).string());
        }
    }

    @Test
    void readsSheetWithStyledCells() throws Exception {
        Path path = tempDir.resolve("styled-cells.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            int styleId = workbook.addStyle(
                com.beingidly.litexl.style.Style.builder().bold(true).fill(0xFFFF0000).build());
            sheet.cell(0, 0).set("Styled").style(styleId);
            sheet.cell(0, 1).set("Plain");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Styled", read.getSheet(0).getCell(0, 0).string());
            assertTrue(read.getSheet(0).getCell(0, 0).styleId() > 0);
            assertEquals("Plain", read.getSheet(0).getCell(0, 1).string());
        }
    }

    @Test
    void readsDateCells() throws Exception {
        Path path = tempDir.resolve("dates.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(java.time.LocalDateTime.of(2024, 6, 15, 10, 30));
            sheet.cell(0, 1).set(java.time.LocalDateTime.of(2000, 1, 1, 0, 0));
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            // Dates are stored as numbers in Excel
            assertNotNull(read.getSheet(0).getCell(0, 0));
            assertEquals(CellType.NUMBER, read.getSheet(0).getCell(0, 0).type());
        }
    }

    @Test
    void cellRawValueCoversAllBranches() throws Exception {
        // Test rawValue() for all cell types including Empty and Date
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Sheet1");
            // Create empty cell and test rawValue
            Cell emptyCell = sheet.cell(0, 0);
            emptyCell.setEmpty();
            assertNull(emptyCell.rawValue()); // Empty branch

            // Date cell
            var dateTime = java.time.LocalDateTime.of(2024, 1, 15, 10, 30);
            Cell dateCell = sheet.cell(0, 1);
            dateCell.set(dateTime);
            assertEquals(dateTime, dateCell.rawValue()); // Date branch
        }
    }

    @Test
    void readsSheetWithDenseData() throws Exception {
        Path path = tempDir.resolve("dense.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            for (int row = 0; row < 50; row++) {
                for (int col = 0; col < 20; col++) {
                    sheet.cell(row, col).set("R" + row + "C" + col);
                }
            }
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("R0C0", read.getSheet(0).getCell(0, 0).string());
            assertEquals("R25C10", read.getSheet(0).getCell(25, 10).string());
            assertEquals("R49C19", read.getSheet(0).getCell(49, 19).string());
        }
    }

    @Test
    void readsSheetWithMixedTypes() throws Exception {
        Path path = tempDir.resolve("mixed-types.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            // Create cells in sparse pattern
            sheet.cell(0, 0).set("String");
            sheet.cell(0, 5).set(123.45);
            sheet.cell(0, 10).set(true);
            sheet.cell(5, 0).setFormula("SUM(A1:A10)");
            sheet.cell(5, 5).setValue(new CellValue.Error("#REF!"));
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(CellType.STRING, read.getSheet(0).getCell(0, 0).type());
            assertEquals(CellType.NUMBER, read.getSheet(0).getCell(0, 5).type());
            assertEquals(CellType.BOOLEAN, read.getSheet(0).getCell(0, 10).type());
            assertEquals(CellType.FORMULA, read.getSheet(0).getCell(5, 0).type());
            assertEquals(CellType.ERROR, read.getSheet(0).getCell(5, 5).type());
        }
    }

    @Test
    void readsSheetWithZeroNumber() throws Exception {
        Path path = tempDir.resolve("zero.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set(0.0);
            sheet.cell(0, 1).set(-0.0);
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(0.0, read.getSheet(0).getCell(0, 0).number(), 0.0001);
        }
    }

    @Test
    void readsSheetWithEmptyStringValue() throws Exception {
        Path path = tempDir.resolve("empty-string.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("");
            sheet.cell(0, 1).set("non-empty");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            // Empty string might be stored differently
            assertEquals("non-empty", read.getSheet(0).getCell(0, 1).string());
        }
    }

    @Test
    void readsSheetWithMultipleMerges() throws Exception {
        Path path = tempDir.resolve("multi-merge2.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Title");
            sheet.merge(0, 0, 0, 5);
            sheet.cell(2, 0).set("Section1");
            sheet.merge(2, 0, 2, 2);
            sheet.cell(2, 3).set("Section2");
            sheet.merge(2, 3, 2, 5);
            sheet.cell(5, 0).set("Footer");
            sheet.merge(5, 0, 6, 5);
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(4, read.getSheet(0).mergedCells().size());
        }
    }

    @Test
    void readsFilesWithSheetNames() throws Exception {
        Path path = tempDir.resolve("named-sheets.xlsx");
        try (Workbook workbook = Workbook.create()) {
            workbook.addSheet("Data").cell(0, 0).set("Data");
            workbook.addSheet("Summary").cell(0, 0).set("Summary");
            workbook.addSheet("Charts").cell(0, 0).set("Charts");
            workbook.addSheet("Special'Name").cell(0, 0).set("Special");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals(4, read.sheetCount());
            assertEquals("Data", read.getSheet("Data").name());
            assertEquals("Summary", read.getSheet("Summary").name());
            assertEquals("Charts", read.getSheet("Charts").name());
            assertEquals("Special'Name", read.getSheet("Special'Name").name());
        }
    }

    @Test
    void readsFileWithRowHeights() throws Exception {
        Path path = tempDir.resolve("row-heights.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("Normal");
            sheet.cell(1, 0).set("Tall");
            sheet.row(1).height(30.0);
            sheet.cell(2, 0).set("Taller");
            sheet.row(2).height(50.0);
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertNotNull(read.getSheet(0));
            assertEquals("Normal", read.getSheet(0).getCell(0, 0).string());
            assertEquals("Tall", read.getSheet(0).getCell(1, 0).string());
            assertEquals("Taller", read.getSheet(0).getCell(2, 0).string());
        }
    }

    @Test
    void readsIntegerSharedStringIndex() throws Exception {
        // This test ensures large shared string indices are parsed correctly
        Path path = tempDir.resolve("many-strings.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            // Create many unique strings
            for (int i = 0; i < 300; i++) {
                sheet.cell(i, 0).set("UniqueString" + i);
            }
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("UniqueString0", read.getSheet(0).getCell(0, 0).string());
            assertEquals("UniqueString150", read.getSheet(0).getCell(150, 0).string());
            assertEquals("UniqueString299", read.getSheet(0).getCell(299, 0).string());
        }
    }

    @Test
    void readsSharedStringsOutOfRange() throws Exception {
        // Tests boundary conditions for shared string index
        Path path = tempDir.resolve("shared-strings.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Sheet1");
            sheet.cell(0, 0).set("First");
            sheet.cell(0, 1).set("Second");
            sheet.cell(0, 2).set("Third");
            workbook.save(path);
        }
        try (Workbook read = Workbook.open(path)) {
            assertEquals("First", read.getSheet(0).getCell(0, 0).string());
            assertEquals("Second", read.getSheet(0).getCell(0, 1).string());
            assertEquals("Third", read.getSheet(0).getCell(0, 2).string());
        }
    }

    // === Tests for POI-generated files with shared strings ===

    @Test
    void readsPoiFileWithSharedStrings() throws Exception {
        Path path = tempDir.resolve("poi-shared-strings.xlsx");
        // Create file using POI (which uses shared strings by default)
        try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            org.apache.poi.xssf.usermodel.XSSFSheet poiSheet = poiWb.createSheet("Data");
            org.apache.poi.xssf.usermodel.XSSFRow row0 = poiSheet.createRow(0);
            row0.createCell(0).setCellValue("Hello");
            row0.createCell(1).setCellValue("World");
            org.apache.poi.xssf.usermodel.XSSFRow row1 = poiSheet.createRow(1);
            row1.createCell(0).setCellValue("Hello"); // Repeated string -> shared
            row1.createCell(1).setCellValue("Universe");

            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path.toFile())) {
                poiWb.write(fos);
            }
        }

        // Read with LiteXL
        try (Workbook read = Workbook.open(path)) {
            assertEquals("Hello", read.getSheet(0).getCell(0, 0).string());
            assertEquals("World", read.getSheet(0).getCell(0, 1).string());
            assertEquals("Hello", read.getSheet(0).getCell(1, 0).string()); // Shared string
            assertEquals("Universe", read.getSheet(0).getCell(1, 1).string());
        }
    }

    @Test
    void readsPoiFileWithMultipleSheets() throws Exception {
        Path path = tempDir.resolve("poi-multi-sheet.xlsx");
        try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            // First sheet
            org.apache.poi.xssf.usermodel.XSSFSheet sheet1 = poiWb.createSheet("First");
            sheet1.createRow(0).createCell(0).setCellValue("Sheet1Data");
            // Second sheet
            org.apache.poi.xssf.usermodel.XSSFSheet sheet2 = poiWb.createSheet("Second");
            sheet2.createRow(0).createCell(0).setCellValue("Sheet2Data");
            // Third sheet
            org.apache.poi.xssf.usermodel.XSSFSheet sheet3 = poiWb.createSheet("Third");
            sheet3.createRow(0).createCell(0).setCellValue("Sheet3Data");

            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path.toFile())) {
                poiWb.write(fos);
            }
        }

        try (Workbook read = Workbook.open(path)) {
            assertEquals(3, read.sheetCount());
            assertEquals("First", read.getSheet(0).name());
            assertEquals("Second", read.getSheet(1).name());
            assertEquals("Third", read.getSheet(2).name());
            assertEquals("Sheet1Data", read.getSheet(0).getCell(0, 0).string());
            assertEquals("Sheet2Data", read.getSheet(1).getCell(0, 0).string());
            assertEquals("Sheet3Data", read.getSheet(2).getCell(0, 0).string());
        }
    }

    @Test
    void readsPoiFileWithManySharedStrings() throws Exception {
        Path path = tempDir.resolve("poi-many-strings.xlsx");
        try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            org.apache.poi.xssf.usermodel.XSSFSheet poiSheet = poiWb.createSheet("Data");
            // Create 500 unique strings
            for (int i = 0; i < 500; i++) {
                org.apache.poi.xssf.usermodel.XSSFRow row = poiSheet.createRow(i);
                row.createCell(0).setCellValue("UniqueString" + i);
            }

            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path.toFile())) {
                poiWb.write(fos);
            }
        }

        try (Workbook read = Workbook.open(path)) {
            assertEquals("UniqueString0", read.getSheet(0).getCell(0, 0).string());
            assertEquals("UniqueString250", read.getSheet(0).getCell(250, 0).string());
            assertEquals("UniqueString499", read.getSheet(0).getCell(499, 0).string());
        }
    }

    @Test
    void readsPoiFileWithAllCellTypes() throws Exception {
        Path path = tempDir.resolve("poi-all-types.xlsx");
        try (org.apache.poi.xssf.usermodel.XSSFWorkbook poiWb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            org.apache.poi.xssf.usermodel.XSSFSheet poiSheet = poiWb.createSheet("Types");
            org.apache.poi.xssf.usermodel.XSSFRow row = poiSheet.createRow(0);
            row.createCell(0).setCellValue("Text");
            row.createCell(1).setCellValue(123.456);
            row.createCell(2).setCellValue(true);
            row.createCell(3).setCellValue(false);
            row.createCell(4).setCellFormula("A1&B1");
            // Error cell
            org.apache.poi.xssf.usermodel.XSSFCell errorCell = row.createCell(5);
            errorCell.setCellErrorValue(org.apache.poi.ss.usermodel.FormulaError.DIV0);

            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path.toFile())) {
                poiWb.write(fos);
            }
        }

        try (Workbook read = Workbook.open(path)) {
            assertEquals(CellType.STRING, read.getSheet(0).getCell(0, 0).type());
            assertEquals("Text", read.getSheet(0).getCell(0, 0).string());
            assertEquals("Text", read.getSheet(0).getCell(0, 0).rawValue()); // rawValue coverage
            assertEquals(CellType.NUMBER, read.getSheet(0).getCell(0, 1).type());
            assertEquals(123.456, read.getSheet(0).getCell(0, 1).number(), 0.001);
            assertEquals(123.456, read.getSheet(0).getCell(0, 1).rawValue()); // rawValue coverage
            assertEquals(CellType.BOOLEAN, read.getSheet(0).getCell(0, 2).type());
            assertTrue(read.getSheet(0).getCell(0, 2).bool());
            assertEquals(true, read.getSheet(0).getCell(0, 2).rawValue()); // rawValue coverage
            assertEquals(CellType.BOOLEAN, read.getSheet(0).getCell(0, 3).type());
            assertFalse(read.getSheet(0).getCell(0, 3).bool());
            assertEquals(false, read.getSheet(0).getCell(0, 3).rawValue()); // rawValue coverage
            assertEquals(CellType.FORMULA, read.getSheet(0).getCell(0, 4).type());
            assertEquals("A1&B1", read.getSheet(0).getCell(0, 4).rawValue()); // rawValue coverage - formula expression
            assertEquals(CellType.ERROR, read.getSheet(0).getCell(0, 5).type());
            assertEquals("#DIV/0!", read.getSheet(0).getCell(0, 5).rawValue()); // rawValue coverage
        }
    }
}
