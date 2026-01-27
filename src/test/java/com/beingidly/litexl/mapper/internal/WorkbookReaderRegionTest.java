package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.*;
import com.beingidly.litexl.mapper.*;
import org.junit.jupiter.api.Test;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class WorkbookReaderRegionTest {

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

    @LitexlSheet
    record MultiRegionSheet(
        List<Employee> employees,
        List<Department> departments
    ) {}

    @LitexlWorkbook
    record MultiRegionWorkbook(
        @LitexlSheet(name = "Data", regionDetection = RegionDetection.AUTO)
        MultiRegionSheet data
    ) {}

    @Test
    void readMultipleRegions() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            // Employee region
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Salary");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(50000.0);
            sheet.cell(2, 0).set("Bob");
            sheet.cell(2, 1).set(60000.0);
            // Department region
            sheet.cell(4, 0).set("DeptName");
            sheet.cell(4, 1).set("Location");
            sheet.cell(5, 0).set("Engineering");
            sheet.cell(5, 1).set("Building A");

            var reader = new WorkbookReader(MapperConfig.defaults());
            var result = reader.read(wb, MultiRegionWorkbook.class);

            assertNotNull(result.data());
            assertEquals(2, result.data().employees().size());
            assertEquals("Alice", result.data().employees().get(0).name());
            assertEquals(1, result.data().departments().size());
            assertEquals("Engineering", result.data().departments().get(0).name());
        }
    }
}
