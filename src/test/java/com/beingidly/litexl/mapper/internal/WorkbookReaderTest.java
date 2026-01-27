package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.*;
import com.beingidly.litexl.mapper.*;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class WorkbookReaderTest {

    @LitexlRow
    record Person(
        @LitexlColumn(index = 0) String name,
        @LitexlColumn(index = 1) int age
    ) {}

    @LitexlWorkbook
    record SimpleWorkbook(
        @LitexlSheet(name = "People", headerRow = 0, dataStartRow = 1)
        List<Person> people
    ) {}

    @Test
    void readSimpleWorkbook() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("People");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Age");
            sheet.cell(1, 0).set("Alice");
            sheet.cell(1, 1).set(30.0);
            sheet.cell(2, 0).set("Bob");
            sheet.cell(2, 1).set(25.0);

            var reader = new WorkbookReader(MapperConfig.defaults());
            var result = reader.read(wb, SimpleWorkbook.class);

            assertNotNull(result);
            assertEquals(2, result.people().size());
            assertEquals("Alice", result.people().get(0).name());
            assertEquals(30, result.people().get(0).age());
            assertEquals("Bob", result.people().get(1).name());
            assertEquals(25, result.people().get(1).age());
        }
    }

    @LitexlWorkbook
    record IndexWorkbook(
        @LitexlSheet(index = 0, dataStartRow = 0)
        List<Person> people
    ) {}

    @Test
    void readBySheetIndex() {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.addSheet("Data");
            sheet.cell(0, 0).set("Charlie");
            sheet.cell(0, 1).set(40.0);

            var reader = new WorkbookReader(MapperConfig.defaults());
            var result = reader.read(wb, IndexWorkbook.class);

            assertEquals(1, result.people().size());
            assertEquals("Charlie", result.people().get(0).name());
        }
    }

    @Test
    void throwsWhenNotAnnotated() {
        record NotAnnotated(String name) {}

        try (Workbook wb = Workbook.create()) {
            var reader = new WorkbookReader(MapperConfig.defaults());
            assertThrows(LitexlMapperException.class, () -> reader.read(wb, NotAnnotated.class));
        }
    }

    @Nested
    class AutoRegionDetectionTests {

        @LitexlRow
        record Product(
            @LitexlColumn(header = "Name") String name,
            @LitexlColumn(header = "Price") int price
        ) {}

        @LitexlWorkbook
        record ProductWorkbook(
            @LitexlSheet(regionDetection = RegionDetection.AUTO)
            List<Product> products
        ) {}

        @Test
        void readListWithAutoRegionDetection() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Sheet1");
                // Title at A1
                sheet.cell(0, 0).set("Product List");
                // Empty row at 1
                // Headers at row 2
                sheet.cell(2, 0).set("Name");
                sheet.cell(2, 1).set("Price");
                // Data starting at row 3
                sheet.cell(3, 0).set("Apple");
                sheet.cell(3, 1).set(100.0);
                sheet.cell(4, 0).set("Banana");
                sheet.cell(4, 1).set(50.0);

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, ProductWorkbook.class);

                assertNotNull(result);
                assertEquals(2, result.products().size());
                assertEquals("Apple", result.products().get(0).name());
                assertEquals(100, result.products().get(0).price());
                assertEquals("Banana", result.products().get(1).name());
                assertEquals(50, result.products().get(1).price());
            }
        }

        @Test
        void readListWithAutoRegionDetection_emptyPrimitiveField() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Sheet1");
                // Headers at row 0
                sheet.cell(0, 0).set("Name");
                sheet.cell(0, 1).set("Price");
                // Data with missing price (empty cell)
                sheet.cell(1, 0).set("Apple");
                // cell(1, 1) is intentionally empty

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, ProductWorkbook.class);

                assertNotNull(result);
                assertEquals(1, result.products().size());
                assertEquals("Apple", result.products().get(0).name());
                assertEquals(0, result.products().get(0).price());  // default value for int
            }
        }

        @Test
        void readListWithAutoRegionDetection_noMatchingHeaders() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Sheet1");
                // Headers that don't match
                sheet.cell(0, 0).set("Column1");
                sheet.cell(0, 1).set("Column2");

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, ProductWorkbook.class);

                assertNotNull(result);
                assertTrue(result.products().isEmpty());
            }
        }
    }

    @Nested
    class ClassBasedWorkbookTests {

        @LitexlRow
        static class Employee {
            @LitexlColumn(index = 0) String name;
            @LitexlColumn(index = 1) int age;
            @LitexlColumn(index = 2) double salary;

            public Employee() {}

            public String getName() { return name; }
            public int getAge() { return age; }
            public double getSalary() { return salary; }
        }

        @LitexlWorkbook
        static class EmployeeWorkbook {
            @LitexlSheet(name = "Employees", headerRow = 0, dataStartRow = 1)
            List<Employee> employees;

            public EmployeeWorkbook() {}

            public List<Employee> getEmployees() { return employees; }
        }

        @Test
        void readClassBasedWorkbook() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Employees");
                // Header row
                sheet.cell(0, 0).set("Name");
                sheet.cell(0, 1).set("Age");
                sheet.cell(0, 2).set("Salary");
                // Data rows
                sheet.cell(1, 0).set("Alice");
                sheet.cell(1, 1).set(30.0);
                sheet.cell(1, 2).set(50000.0);
                sheet.cell(2, 0).set("Bob");
                sheet.cell(2, 1).set(25.0);
                sheet.cell(2, 2).set(45000.0);

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, EmployeeWorkbook.class);

                assertNotNull(result);
                assertEquals(2, result.getEmployees().size());
                assertEquals("Alice", result.getEmployees().get(0).getName());
                assertEquals(30, result.getEmployees().get(0).getAge());
                assertEquals(50000.0, result.getEmployees().get(0).getSalary());
                assertEquals("Bob", result.getEmployees().get(1).getName());
                assertEquals(25, result.getEmployees().get(1).getAge());
                assertEquals(45000.0, result.getEmployees().get(1).getSalary());
            }
        }

        @LitexlRow
        static class Item {
            @LitexlColumn(index = 0, header = "Item Name") String itemName;
            @LitexlColumn(index = 1, header = "Quantity") int quantity;

            public Item() {}

            public String getItemName() { return itemName; }
            public int getQuantity() { return quantity; }
        }

        @LitexlWorkbook
        static class ItemWorkbook {
            @LitexlSheet(regionDetection = RegionDetection.AUTO)
            List<Item> items;

            public ItemWorkbook() {}

            public List<Item> getItems() { return items; }
        }

        @Test
        void readClassBasedWorkbookWithAutoRegionDetection() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Sheet1");
                // Title row
                sheet.cell(0, 0).set("Inventory Report");
                // Empty row
                // Headers at row 2
                sheet.cell(2, 0).set("Item Name");
                sheet.cell(2, 1).set("Quantity");
                // Data
                sheet.cell(3, 0).set("Widget");
                sheet.cell(3, 1).set(100.0);
                sheet.cell(4, 0).set("Gadget");
                sheet.cell(4, 1).set(50.0);

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, ItemWorkbook.class);

                assertNotNull(result);
                assertEquals(2, result.getItems().size());
                assertEquals("Widget", result.getItems().get(0).getItemName());
                assertEquals(100, result.getItems().get(0).getQuantity());
                assertEquals("Gadget", result.getItems().get(1).getItemName());
                assertEquals(50, result.getItems().get(1).getQuantity());
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
        void readLitexlCellAtWorkbookLevel() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Sheet1");
                sheet.cell(0, 0).set("Monthly Report");
                sheet.cell(1, 0).set("John Doe");
                sheet.cell(2, 0).set(1.5);

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, ReportWorkbook.class);

                assertNotNull(result);
                assertEquals("Monthly Report", result.title());
                assertEquals("John Doe", result.author());
                assertEquals(1.5, result.version());
            }
        }

        @LitexlWorkbook
        static class ConfigWorkbook {
            @LitexlCell(row = 0, column = 0) String appName;
            @LitexlCell(row = 0, column = 1) int maxConnections;
            @LitexlCell(row = 1, column = 0) boolean debugMode;

            public ConfigWorkbook() {}

            public String getAppName() { return appName; }
            public int getMaxConnections() { return maxConnections; }
            public boolean isDebugMode() { return debugMode; }
        }

        @Test
        void readLitexlCellAtWorkbookLevelWithClassBasedType() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Config");
                sheet.cell(0, 0).set("MyApp");
                sheet.cell(0, 1).set(100.0);
                sheet.cell(1, 0).set(true);

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, ConfigWorkbook.class);

                assertNotNull(result);
                assertEquals("MyApp", result.getAppName());
                assertEquals(100, result.getMaxConnections());
                assertTrue(result.isDebugMode());
            }
        }

        @LitexlRow
        record DataRow(
            @LitexlColumn(index = 0) String col1,
            @LitexlColumn(index = 1) int col2
        ) {}

        @LitexlWorkbook
        record MixedWorkbook(
            @LitexlCell(row = 0, column = 0) String header,
            @LitexlSheet(name = "Data", headerRow = 2, dataStartRow = 3)
            List<DataRow> rows
        ) {}

        @Test
        void readMixedCellAndSheetAnnotations() {
            try (Workbook wb = Workbook.create()) {
                // First sheet has header cell
                Sheet sheet1 = wb.addSheet("Sheet1");
                sheet1.cell(0, 0).set("Report Header");

                // Second sheet has data
                Sheet sheet2 = wb.addSheet("Data");
                sheet2.cell(2, 0).set("Col1");
                sheet2.cell(2, 1).set("Col2");
                sheet2.cell(3, 0).set("A");
                sheet2.cell(3, 1).set(1.0);
                sheet2.cell(4, 0).set("B");
                sheet2.cell(4, 1).set(2.0);

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, MixedWorkbook.class);

                assertNotNull(result);
                assertEquals("Report Header", result.header());
                assertEquals(2, result.rows().size());
                assertEquals("A", result.rows().get(0).col1());
                assertEquals(1, result.rows().get(0).col2());
                assertEquals("B", result.rows().get(1).col1());
                assertEquals(2, result.rows().get(1).col2());
            }
        }
    }

    @Nested
    class EdgeCaseTests {

        @LitexlRow
        record VariousTypes(
            @LitexlColumn(index = 0) String text,
            @LitexlColumn(index = 1) int intVal,
            @LitexlColumn(index = 2) double doubleVal,
            @LitexlColumn(index = 3) boolean boolVal,
            @LitexlColumn(index = 4) long longVal
        ) {}

        @LitexlWorkbook
        record TypesWorkbook(
            @LitexlSheet(name = "Types", dataStartRow = 1)
            List<VariousTypes> types
        ) {}

        @Test
        void readVariousTypesFromCells() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Types");
                sheet.cell(0, 0).set("Text");
                sheet.cell(0, 1).set("Int");
                sheet.cell(0, 2).set("Double");
                sheet.cell(0, 3).set("Bool");
                sheet.cell(0, 4).set("Long");
                // Data row
                sheet.cell(1, 0).set("Hello");
                sheet.cell(1, 1).set(42.0);
                sheet.cell(1, 2).set(3.14);
                sheet.cell(1, 3).set(true);
                sheet.cell(1, 4).set(9999999999L);

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, TypesWorkbook.class);

                assertNotNull(result);
                assertEquals(1, result.types().size());
                var row = result.types().get(0);
                assertEquals("Hello", row.text());
                assertEquals(42, row.intVal());
                assertEquals(3.14, row.doubleVal(), 0.01);
                assertTrue(row.boolVal());
                assertEquals(9999999999L, row.longVal());
            }
        }

        @LitexlRow
        record NullableRow(
            @LitexlColumn(index = 0) String text,
            @LitexlColumn(index = 1) Integer intVal
        ) {}

        @LitexlWorkbook
        record NullableWorkbook(
            @LitexlSheet(name = "Sheet1", dataStartRow = 0)
            List<NullableRow> rows
        ) {}

        @Test
        void readNullableFieldsFromEmptyCells() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Sheet1");
                // First row has all values
                sheet.cell(0, 0).set("Text1");
                sheet.cell(0, 1).set(10.0);
                // Second row has empty cells - only text
                sheet.cell(1, 0).set("Text2");
                // cell(1, 1) is empty

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, NullableWorkbook.class);

                assertNotNull(result);
                assertEquals(2, result.rows().size());
                assertEquals("Text1", result.rows().get(0).text());
                assertEquals(10, result.rows().get(0).intVal());
                assertEquals("Text2", result.rows().get(1).text());
                assertNull(result.rows().get(1).intVal());
            }
        }

        @LitexlRow
        record HeaderlessRow(
            @LitexlColumn(index = 0) String col0,
            @LitexlColumn(index = 1) String col1
        ) {}

        @LitexlWorkbook
        record HeaderlessWorkbook(
            @LitexlSheet(index = 0, dataStartRow = 0)
            List<HeaderlessRow> rows
        ) {}

        @Test
        void readSheetWithoutHeaders() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Data");
                // Data starts at row 0, no headers
                sheet.cell(0, 0).set("A1");
                sheet.cell(0, 1).set("B1");
                sheet.cell(1, 0).set("A2");
                sheet.cell(1, 1).set("B2");

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, HeaderlessWorkbook.class);

                assertNotNull(result);
                assertEquals(2, result.rows().size());
                assertEquals("A1", result.rows().get(0).col0());
                assertEquals("B1", result.rows().get(0).col1());
                assertEquals("A2", result.rows().get(1).col0());
                assertEquals("B2", result.rows().get(1).col1());
            }
        }

        @LitexlRow
        record SparseRow(
            @LitexlColumn(index = 0) String col0,
            @LitexlColumn(index = 5) String col5,
            @LitexlColumn(index = 10) String col10
        ) {}

        @LitexlWorkbook
        record SparseWorkbook(
            @LitexlSheet(name = "Sparse", dataStartRow = 0)
            List<SparseRow> rows
        ) {}

        @Test
        void readSparseColumns() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet = wb.addSheet("Sparse");
                sheet.cell(0, 0).set("First");
                // Columns 1-4 are empty
                sheet.cell(0, 5).set("Middle");
                // Columns 6-9 are empty
                sheet.cell(0, 10).set("Last");

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, SparseWorkbook.class);

                assertNotNull(result);
                assertEquals(1, result.rows().size());
                assertEquals("First", result.rows().get(0).col0());
                assertEquals("Middle", result.rows().get(0).col5());
                assertEquals("Last", result.rows().get(0).col10());
            }
        }

        @LitexlWorkbook
        record MultiSheetWorkbook(
            @LitexlSheet(name = "People", dataStartRow = 0)
            List<Person> people,
            @LitexlSheet(name = "Items", dataStartRow = 0)
            List<HeaderlessRow> items
        ) {}

        @Test
        void readMultipleSheets() {
            try (Workbook wb = Workbook.create()) {
                Sheet sheet1 = wb.addSheet("People");
                sheet1.cell(0, 0).set("Alice");
                sheet1.cell(0, 1).set(25.0);
                sheet1.cell(1, 0).set("Bob");
                sheet1.cell(1, 1).set(30.0);

                Sheet sheet2 = wb.addSheet("Items");
                sheet2.cell(0, 0).set("Item1");
                sheet2.cell(0, 1).set("Desc1");

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, MultiSheetWorkbook.class);

                assertNotNull(result);
                assertEquals(2, result.people().size());
                assertEquals("Alice", result.people().get(0).name());
                assertEquals(1, result.items().size());
                assertEquals("Item1", result.items().get(0).col0());
            }
        }

        @LitexlWorkbook
        record EmptySheetWorkbook(
            @LitexlSheet(name = "Empty", dataStartRow = 0)
            List<Person> people
        ) {}

        @Test
        void readEmptySheet() {
            try (Workbook wb = Workbook.create()) {
                wb.addSheet("Empty"); // Empty sheet

                var reader = new WorkbookReader(MapperConfig.defaults());
                var result = reader.read(wb, EmptySheetWorkbook.class);

                assertNotNull(result);
                assertTrue(result.people().isEmpty());
            }
        }
    }
}
