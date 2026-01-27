package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.*;
import com.beingidly.litexl.mapper.*;
import org.junit.jupiter.api.Test;
import java.util.List;
import java.util.Set;
import static org.junit.jupiter.api.Assertions.*;

class RegionDetectorTest {

    @LitexlRow
    record Employee(
        @LitexlColumn(header = "Name") String name,
        @LitexlColumn(header = "Salary") double salary
    ) {}

    @LitexlRow
    record Department(
        @LitexlColumn(header = "DeptName") String name,
        @LitexlColumn(header = "Location") String location
    ) {}

    @LitexlRow
    record MixedRecord(
        @LitexlColumn(index = 0) String noHeader,
        @LitexlColumn(header = "Title") String withHeader,
        @LitexlColumn(header = "") String emptyHeader
    ) {}

    @LitexlRow
    record NoHeaderRecord(
        @LitexlColumn(index = 0) String col1,
        @LitexlColumn(index = 1) String col2
    ) {}

    @Test
    void detectSingleRegion() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(50000.0);

            var headers = Set.of("Name", "Salary");
            var region = RegionDetector.detectRegion(sheet, headers, 0);

            assertNotNull(region);
            assertEquals(0, region.headerRow());
            assertEquals(1, region.dataStartRow());
        }
    }

    @Test
    void detectMultipleRegions() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            // Employee region
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(50000.0);
            // Department region
            sheet.cell(3, 0).set("DeptName");
            sheet.cell(3, 1).set("Location");
            sheet.cell(4, 0).set("Engineering");
            sheet.cell(4, 1).set("Building A");

            var employeeHeaders = Set.of("Name", "Salary");
            var deptHeaders = Set.of("DeptName", "Location");

            var empRegion = RegionDetector.detectRegion(sheet, employeeHeaders, 0);
            var deptRegion = RegionDetector.detectRegion(sheet, deptHeaders, 0);

            assertNotNull(empRegion);
            assertEquals(0, empRegion.headerRow());

            assertNotNull(deptRegion);
            assertEquals(3, deptRegion.headerRow());
        }
    }

    @Test
    void detectRegionEndRow() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(50000.0);
            sheet.cell(2, 0).set("Bob");
            sheet.cell(2, 1).set(60000.0);
            // Empty row 3
            sheet.cell(4, 0).set("DeptName");
            sheet.cell(4, 1).set("Location");

            var headers = Set.of("Name", "Salary");
            var nextHeaders = Set.of("DeptName", "Location");
            var region = RegionDetector.detectRegionWithEnd(sheet, headers, nextHeaders, 0);

            assertNotNull(region);
            assertEquals(0, region.headerRow());
            assertEquals(1, region.dataStartRow());
            assertEquals(2, region.dataEndRow());
        }
    }

    @Test
    void noRegionFound() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Other");
            sheet.cell(0, 1).set("Headers");

            var headers = Set.of("Name", "Salary");
            var region = RegionDetector.detectRegion(sheet, headers, 0);

            assertNull(region);
        }
    }

    // ==================== Empty Sheet Tests ====================

    @Test
    void detectRegionOnEmptySheet() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Empty");

            var headers = Set.of("Name", "Salary");
            var region = RegionDetector.detectRegion(sheet, headers, 0);

            assertNull(region);
        }
    }

    @Test
    void detectRegionWithEndOnEmptySheet() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Empty");

            var headers = Set.of("Name", "Salary");
            var nextHeaders = Set.of("DeptName", "Location");
            var region = RegionDetector.detectRegionWithEnd(sheet, headers, nextHeaders, 0);

            assertNull(region);
        }
    }

    // ==================== Single Cell Tests ====================

    @Test
    void detectRegionWithSingleCellHeader() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");

            var headers = Set.of("Name");
            var region = RegionDetector.detectRegion(sheet, headers, 0);

            assertNotNull(region);
            assertEquals(0, region.headerRow());
            assertEquals(1, region.dataStartRow());
        }
    }

    @Test
    void detectRegionWithSingleCellData() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");
            sheet.cell(1, 0).set("Alice");

            var headers = Set.of("Name");
            var nextHeaders = Set.of("Other");
            var region = RegionDetector.detectRegionWithEnd(sheet, headers, nextHeaders, 0);

            assertNotNull(region);
            assertEquals(0, region.headerRow());
            assertEquals(1, region.dataStartRow());
            assertEquals(1, region.dataEndRow());
        }
    }

    // ==================== Sparse Data Tests ====================

    @Test
    void detectRegionWithSparseHeaders() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            // Headers with gap in columns
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 5).set("Salary");  // Gap at columns 1-4
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 5).set(50000.0);

            var headers = Set.of("Name", "Salary");
            var region = RegionDetector.detectRegion(sheet, headers, 0);

            assertNotNull(region);
            assertEquals(0, region.headerRow());
            assertEquals(1, region.dataStartRow());
        }
    }

    @Test
    void detectRegionWithEmptyRowsInData() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(50000.0);
            // Empty row 2 (skipped)
            sheet.cell(3, 0).set("Bob");
            sheet.cell(3, 1).set(60000.0);
            // Next section starts at row 5
            sheet.cell(5, 0).set("DeptName");
            sheet.cell(5, 1).set("Location");

            var headers = Set.of("Name", "Salary");
            var nextHeaders = Set.of("DeptName", "Location");
            var region = RegionDetector.detectRegionWithEnd(sheet, headers, nextHeaders, 0);

            assertNotNull(region);
            assertEquals(0, region.headerRow());
            assertEquals(1, region.dataStartRow());
            // End row should be 3 (last data row before next headers)
            assertEquals(3, region.dataEndRow());
        }
    }

    @Test
    void detectRegionWithNonStringCells() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set(12345.0);  // Number, not string
            sheet.cell(0, 2).set("Salary");

            // Should only match string headers
            var headers = Set.of("Name", "Salary");
            var region = RegionDetector.detectRegion(sheet, headers, 0);

            assertNotNull(region);
            assertEquals(0, region.headerRow());
        }
    }

    // ==================== Data With Headers Tests ====================

    @Test
    void detectRegionWithExtraHeaders() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            // Row has more headers than required
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(0, 2).set("Department");  // Extra header
            sheet.cell(0, 3).set("Age");         // Extra header

            var headers = Set.of("Name", "Salary");
            var region = RegionDetector.detectRegion(sheet, headers, 0);

            // Should still find the region (containsAll check)
            assertNotNull(region);
            assertEquals(0, region.headerRow());
        }
    }

    @Test
    void detectRegionWithPartialHeaders() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            // Row has fewer headers than required
            sheet.cell(0, 0).set("Name");
            // Missing "Salary"

            var headers = Set.of("Name", "Salary");
            var region = RegionDetector.detectRegion(sheet, headers, 0);

            assertNull(region);
        }
    }

    @Test
    void detectRegionWithHeadersAtDifferentRow() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            // Title row
            sheet.cell(0, 0).set("Report Title");
            // Empty row 1
            // Headers at row 2
            sheet.cell(2, 0).set("Name");
            sheet.cell(2, 1).set("Salary");
            sheet.cell(3, 0).set("Alice");
            sheet.cell(3, 1).set(50000.0);

            var headers = Set.of("Name", "Salary");
            var region = RegionDetector.detectRegion(sheet, headers, 0);

            assertNotNull(region);
            assertEquals(2, region.headerRow());
            assertEquals(3, region.dataStartRow());
        }
    }

    // ==================== Region Record Tests ====================

    @Test
    void regionRecordTwoArgConstructor() {
        var region = new RegionDetector.Region(5, 6);

        assertEquals(5, region.headerRow());
        assertEquals(6, region.dataStartRow());
        assertEquals(Integer.MAX_VALUE, region.dataEndRow());
    }

    @Test
    void regionRecordThreeArgConstructor() {
        var region = new RegionDetector.Region(5, 6, 10);

        assertEquals(5, region.headerRow());
        assertEquals(6, region.dataStartRow());
        assertEquals(10, region.dataEndRow());
    }

    @Test
    void regionRecordEquality() {
        var region1 = new RegionDetector.Region(0, 1, 5);
        var region2 = new RegionDetector.Region(0, 1, 5);
        var region3 = new RegionDetector.Region(0, 1, 10);

        assertEquals(region1, region2);
        assertNotEquals(region1, region3);
    }

    // ==================== extractHeaders Tests ====================

    @Test
    void extractHeadersFromEmployee() {
        var headers = RegionDetector.extractHeaders(Employee.class);

        assertEquals(2, headers.size());
        assertTrue(headers.contains("Name"));
        assertTrue(headers.contains("Salary"));
    }

    @Test
    void extractHeadersFromDepartment() {
        var headers = RegionDetector.extractHeaders(Department.class);

        assertEquals(2, headers.size());
        assertTrue(headers.contains("DeptName"));
        assertTrue(headers.contains("Location"));
    }

    @Test
    void extractHeadersWithMixedAnnotations() {
        // MixedRecord has: index-only, header-with-value, empty-header
        var headers = RegionDetector.extractHeaders(MixedRecord.class);

        // Only the one with actual header value should be included
        assertEquals(1, headers.size());
        assertTrue(headers.contains("Title"));
    }

    @Test
    void extractHeadersFromIndexOnlyRecord() {
        // NoHeaderRecord has only index annotations, no headers
        var headers = RegionDetector.extractHeaders(NoHeaderRecord.class);

        assertTrue(headers.isEmpty());
    }

    // ==================== startRow Parameter Tests ====================

    @Test
    void detectRegionWithStartRowSkipsEarlierMatches() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            // First occurrence at row 0
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(50000.0);
            // Second occurrence at row 5
            sheet.cell(5, 0).set("Name");
            sheet.cell(5, 1).set("Salary");
            sheet.cell(6, 0).set("Bob");
            sheet.cell(6, 1).set(60000.0);

            var headers = Set.of("Name", "Salary");

            // Start from row 3 - should skip first occurrence
            var region = RegionDetector.detectRegion(sheet, headers, 3);

            assertNotNull(region);
            assertEquals(5, region.headerRow());
            assertEquals(6, region.dataStartRow());
        }
    }

    @Test
    void detectRegionWithStartRowBeyondData() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");

            var headers = Set.of("Name", "Salary");

            // Start from row 10 - beyond any data
            var region = RegionDetector.detectRegion(sheet, headers, 10);

            assertNull(region);
        }
    }

    @Test
    void detectRegionWithEndWithStartRowSkipsEarlierMatches() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            // First section
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(50000.0);
            // Second section (same headers)
            sheet.cell(5, 0).set("Name");
            sheet.cell(5, 1).set("Salary");
            sheet.cell(6, 0).set("Bob");
            sheet.cell(6, 1).set(60000.0);
            // Next section headers
            sheet.cell(8, 0).set("DeptName");
            sheet.cell(8, 1).set("Location");

            var headers = Set.of("Name", "Salary");
            var nextHeaders = Set.of("DeptName", "Location");

            // Start from row 3 - should detect second occurrence
            var region = RegionDetector.detectRegionWithEnd(sheet, headers, nextHeaders, 3);

            assertNotNull(region);
            assertEquals(5, region.headerRow());
            assertEquals(6, region.dataStartRow());
            assertEquals(6, region.dataEndRow());
        }
    }

    // ==================== detectRegionWithEnd Edge Cases ====================

    @Test
    void detectRegionWithEndNoNextHeaders() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(50000.0);
            sheet.cell(2, 0).set("Bob");
            sheet.cell(2, 1).set(60000.0);

            var headers = Set.of("Name", "Salary");
            var nextHeaders = Set.of("NonExistent");  // Won't be found
            var region = RegionDetector.detectRegionWithEnd(sheet, headers, nextHeaders, 0);

            assertNotNull(region);
            assertEquals(0, region.headerRow());
            assertEquals(1, region.dataStartRow());
            // Should end at last data row
            assertEquals(2, region.dataEndRow());
        }
    }

    @Test
    void detectRegionWithEndImmediateNextHeaders() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            // No data rows - next headers immediately follow
            sheet.cell(1, 0).set("DeptName");
            sheet.cell(1, 1).set("Location");

            var headers = Set.of("Name", "Salary");
            var nextHeaders = Set.of("DeptName", "Location");
            var region = RegionDetector.detectRegionWithEnd(sheet, headers, nextHeaders, 0);

            assertNotNull(region);
            assertEquals(0, region.headerRow());
            assertEquals(1, region.dataStartRow());
            // End row should be 0 (dataStartRow - 1 since no data rows)
            assertEquals(0, region.dataEndRow());
        }
    }

    @Test
    void detectRegionWithEndFirstRegionNotFound() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Other");
            sheet.cell(0, 1).set("Headers");

            var headers = Set.of("Name", "Salary");
            var nextHeaders = Set.of("DeptName", "Location");
            var region = RegionDetector.detectRegionWithEnd(sheet, headers, nextHeaders, 0);

            assertNull(region);
        }
    }

    @Test
    void detectRegionWithMultipleEmptyRowsBeforeNext() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(50000.0);
            // Multiple empty rows (2, 3, 4)
            sheet.cell(5, 0).set("DeptName");
            sheet.cell(5, 1).set("Location");

            var headers = Set.of("Name", "Salary");
            var nextHeaders = Set.of("DeptName", "Location");
            var region = RegionDetector.detectRegionWithEnd(sheet, headers, nextHeaders, 0);

            assertNotNull(region);
            assertEquals(1, region.dataEndRow());
        }
    }
}
