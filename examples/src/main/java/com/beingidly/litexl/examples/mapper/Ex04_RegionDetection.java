package com.beingidly.litexl.examples.mapper;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.mapper.LitexlColumn;
import com.beingidly.litexl.mapper.LitexlMapper;
import com.beingidly.litexl.mapper.LitexlRow;
import com.beingidly.litexl.mapper.LitexlSheet;
import com.beingidly.litexl.mapper.LitexlWorkbook;
import com.beingidly.litexl.mapper.RegionDetection;

import java.nio.file.Path;
import java.util.List;

/**
 * Example: Auto-detecting data regions in Excel files.
 *
 * This example demonstrates:
 * - Using RegionDetection.AUTO to find data automatically
 * - Reading Excel files where data doesn't start at A1
 */
public class Ex04_RegionDetection {

    @LitexlRow
    public record Employee(
        @LitexlColumn(index = 0, header = "ID")
        int id,

        @LitexlColumn(index = 1, header = "Name")
        String name,

        @LitexlColumn(index = 2, header = "Department")
        String department
    ) {}

    // Use AUTO region detection to find the data table automatically
    @LitexlWorkbook
    public record EmployeeReport(
        @LitexlSheet(name = "Employees", regionDetection = RegionDetection.AUTO)
        List<Employee> employees
    ) {}

    public static void main(String[] args) {
        // First, create an Excel file where data doesn't start at A1
        Path samplePath = createSampleFileWithOffset();

        System.out.println("Created sample file with data starting at C3 (not A1)");
        System.out.println();

        // Read using AUTO region detection
        System.out.println("Reading with RegionDetection.AUTO:");

        EmployeeReport report = LitexlMapper.read(samplePath, EmployeeReport.class);

        for (Employee emp : report.employees()) {
            System.out.println("  ID: " + emp.id() + ", Name: " + emp.name() + ", Dept: " + emp.department());
        }
    }

    private static Path createSampleFileWithOffset() {
        Path path = ExampleUtils.tempFile("ex04_region_detection.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Employees");

            // Add a title at A1
            sheet.cell(0, 0).set("Employee Directory");

            // Add some empty space, then data starting at C3 (row 2, col 2)
            int startRow = 2;
            int startCol = 2;

            // Headers
            sheet.cell(startRow, startCol).set("ID");
            sheet.cell(startRow, startCol + 1).set("Name");
            sheet.cell(startRow, startCol + 2).set("Department");

            // Data
            sheet.cell(startRow + 1, startCol).set(1);
            sheet.cell(startRow + 1, startCol + 1).set("Alice");
            sheet.cell(startRow + 1, startCol + 2).set("Engineering");

            sheet.cell(startRow + 2, startCol).set(2);
            sheet.cell(startRow + 2, startCol + 1).set("Bob");
            sheet.cell(startRow + 2, startCol + 2).set("Marketing");

            sheet.cell(startRow + 3, startCol).set(3);
            sheet.cell(startRow + 3, startCol + 1).set("Charlie");
            sheet.cell(startRow + 3, startCol + 2).set("Finance");

            workbook.save(path);
        }

        ExampleUtils.printCreated(path);
        return path;
    }
}
