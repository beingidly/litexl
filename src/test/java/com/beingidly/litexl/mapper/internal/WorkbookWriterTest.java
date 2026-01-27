package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.*;
import com.beingidly.litexl.mapper.*;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class WorkbookWriterTest {

    @LitexlRow
    record Person(
        @LitexlColumn(index = 0, header = "Name") String name,
        @LitexlColumn(index = 1, header = "Age") int age
    ) {}

    @LitexlWorkbook
    record SimpleWorkbook(
        @LitexlSheet(name = "People")
        List<Person> people
    ) {}

    @Test
    void writeSimpleWorkbook() {
        var data = new SimpleWorkbook(List.of(
            new Person("Alice", 30),
            new Person("Bob", 25)
        ));

        try (Workbook wb = Workbook.create()) {
            var writer = new WorkbookWriter(MapperConfig.defaults());
            writer.write(wb, data);

            assertEquals(1, wb.sheetCount());
            var sheet = wb.getSheet("People");
            assertNotNull(sheet);

            // Header row
            assertEquals("Name", sheet.cell(0, 0).string());
            assertEquals("Age", sheet.cell(0, 1).string());

            // Data rows
            assertEquals("Alice", sheet.cell(1, 0).string());
            assertEquals(30.0, sheet.cell(1, 1).number());
            assertEquals("Bob", sheet.cell(2, 0).string());
            assertEquals(25.0, sheet.cell(2, 1).number());
        }
    }

    @LitexlWorkbook
    record OffsetWorkbook(
        @LitexlSheet(name = "Data", headerRow = 2, dataStartRow = 3, dataStartColumn = 1)
        List<Person> people
    ) {}

    @Test
    void writeWithOffset() {
        var data = new OffsetWorkbook(List.of(new Person("Charlie", 35)));

        try (Workbook wb = Workbook.create()) {
            var writer = new WorkbookWriter(MapperConfig.defaults());
            writer.write(wb, data);

            var sheet = wb.getSheet("Data");

            // Header at row 2, column 1 (B3 in Excel)
            assertEquals("Name", sheet.cell(2, 1).string());
            assertEquals("Age", sheet.cell(2, 2).string());

            // Data at row 3, column 1
            assertEquals("Charlie", sheet.cell(3, 1).string());
            assertEquals(35.0, sheet.cell(3, 2).number());
        }
    }

    @Test
    void throwsWhenNotAnnotated() {
        record NotAnnotated(String name) {}

        try (Workbook wb = Workbook.create()) {
            var writer = new WorkbookWriter(MapperConfig.defaults());
            assertThrows(LitexlMapperException.class, () -> writer.write(wb, new NotAnnotated("test")));
        }
    }

    @Nested
    class ClassBasedWorkbookTests {

        @LitexlRow
        static class Employee {
            @LitexlColumn(index = 0, header = "Name") String name;
            @LitexlColumn(index = 1, header = "Age") int age;
            @LitexlColumn(index = 2, header = "Salary") double salary;

            public Employee() {}

            public Employee(String name, int age, double salary) {
                this.name = name;
                this.age = age;
                this.salary = salary;
            }
        }

        @LitexlWorkbook
        static class EmployeeWorkbook {
            @LitexlSheet(name = "Employees")
            List<Employee> employees;

            public EmployeeWorkbook() {}

            public EmployeeWorkbook(List<Employee> employees) {
                this.employees = employees;
            }
        }

        @Test
        void writeClassBasedWorkbook() {
            var data = new EmployeeWorkbook(List.of(
                new Employee("Alice", 30, 50000.0),
                new Employee("Bob", 25, 45000.0)
            ));

            try (Workbook wb = Workbook.create()) {
                var writer = new WorkbookWriter(MapperConfig.defaults());
                writer.write(wb, data);

                assertEquals(1, wb.sheetCount());
                var sheet = wb.getSheet("Employees");
                assertNotNull(sheet);

                // Header row
                assertEquals("Name", sheet.cell(0, 0).string());
                assertEquals("Age", sheet.cell(0, 1).string());
                assertEquals("Salary", sheet.cell(0, 2).string());

                // Data rows
                assertEquals("Alice", sheet.cell(1, 0).string());
                assertEquals(30.0, sheet.cell(1, 1).number());
                assertEquals(50000.0, sheet.cell(1, 2).number());
                assertEquals("Bob", sheet.cell(2, 0).string());
                assertEquals(25.0, sheet.cell(2, 1).number());
                assertEquals(45000.0, sheet.cell(2, 2).number());
            }
        }

        @LitexlWorkbook
        static class OffsetEmployeeWorkbook {
            @LitexlSheet(name = "Data", headerRow = 2, dataStartRow = 3, dataStartColumn = 1)
            List<Employee> employees;

            public OffsetEmployeeWorkbook() {}

            public OffsetEmployeeWorkbook(List<Employee> employees) {
                this.employees = employees;
            }
        }

        @Test
        void writeClassBasedWorkbookWithOffset() {
            var data = new OffsetEmployeeWorkbook(List.of(
                new Employee("Charlie", 35, 60000.0)
            ));

            try (Workbook wb = Workbook.create()) {
                var writer = new WorkbookWriter(MapperConfig.defaults());
                writer.write(wb, data);

                var sheet = wb.getSheet("Data");
                assertNotNull(sheet);

                // Header at row 2, column 1 (B3 in Excel)
                assertEquals("Name", sheet.cell(2, 1).string());
                assertEquals("Age", sheet.cell(2, 2).string());
                assertEquals("Salary", sheet.cell(2, 3).string());

                // Data at row 3, column 1
                assertEquals("Charlie", sheet.cell(3, 1).string());
                assertEquals(35.0, sheet.cell(3, 2).number());
                assertEquals(60000.0, sheet.cell(3, 3).number());
            }
        }
    }

    @Nested
    class LitexlCellAtWorkbookLevelTests {

        @LitexlWorkbook
        record ReportWorkbook(
            @LitexlCell(row = 0, column = 0) String title,
            @LitexlCell(row = 1, column = 0) String author,
            @LitexlCell(row = 2, column = 0) double version
        ) {}

        @Test
        void writeLitexlCellAtWorkbookLevel() {
            var data = new ReportWorkbook("Monthly Report", "John Doe", 1.5);

            try (Workbook wb = Workbook.create()) {
                var writer = new WorkbookWriter(MapperConfig.defaults());
                writer.write(wb, data);

                assertEquals(1, wb.sheetCount());
                var sheet = wb.getSheet(0);
                assertNotNull(sheet);

                assertEquals("Monthly Report", sheet.cell(0, 0).string());
                assertEquals("John Doe", sheet.cell(1, 0).string());
                assertEquals(1.5, sheet.cell(2, 0).number());
            }
        }

        @LitexlWorkbook
        static class ConfigWorkbook {
            @LitexlCell(row = 0, column = 0) String appName;
            @LitexlCell(row = 0, column = 1) int maxConnections;
            @LitexlCell(row = 1, column = 0) boolean debugMode;

            public ConfigWorkbook() {}

            public ConfigWorkbook(String appName, int maxConnections, boolean debugMode) {
                this.appName = appName;
                this.maxConnections = maxConnections;
                this.debugMode = debugMode;
            }
        }

        @Test
        void writeLitexlCellAtWorkbookLevelWithClassBasedType() {
            var data = new ConfigWorkbook("MyApp", 100, true);

            try (Workbook wb = Workbook.create()) {
                var writer = new WorkbookWriter(MapperConfig.defaults());
                writer.write(wb, data);

                assertEquals(1, wb.sheetCount());
                var sheet = wb.getSheet(0);
                assertNotNull(sheet);

                assertEquals("MyApp", sheet.cell(0, 0).string());
                assertEquals(100.0, sheet.cell(0, 1).number());
                assertTrue(sheet.cell(1, 0).bool());
            }
        }

        @LitexlRow
        record DataRow(
            @LitexlColumn(index = 0, header = "Col1") String col1,
            @LitexlColumn(index = 1, header = "Col2") int col2
        ) {}

        @LitexlWorkbook
        record MixedWorkbook(
            @LitexlCell(row = 0, column = 0) String header,
            @LitexlSheet(name = "Data", headerRow = 0, dataStartRow = 1)
            List<DataRow> rows
        ) {}

        @Test
        void writeMixedCellAndSheetAnnotations() {
            var data = new MixedWorkbook(
                "Report Header",
                List.of(new DataRow("A", 1), new DataRow("B", 2))
            );

            try (Workbook wb = Workbook.create()) {
                var writer = new WorkbookWriter(MapperConfig.defaults());
                writer.write(wb, data);

                // Should have 2 sheets: Sheet1 for @LitexlCell and Data for @LitexlSheet
                assertEquals(2, wb.sheetCount());

                // First sheet has the cell data
                var sheet1 = wb.getSheet(0);
                assertNotNull(sheet1);
                assertEquals("Report Header", sheet1.cell(0, 0).string());

                // Data sheet has the list data
                var dataSheet = wb.getSheet("Data");
                assertNotNull(dataSheet);
                assertEquals("Col1", dataSheet.cell(0, 0).string());
                assertEquals("Col2", dataSheet.cell(0, 1).string());
                assertEquals("A", dataSheet.cell(1, 0).string());
                assertEquals(1.0, dataSheet.cell(1, 1).number());
                assertEquals("B", dataSheet.cell(2, 0).string());
                assertEquals(2.0, dataSheet.cell(2, 1).number());
            }
        }
    }
}
