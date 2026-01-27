package com.beingidly.litexl.mapper;

import com.beingidly.litexl.CellValue;
import com.beingidly.litexl.style.Style;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.temporal.ChronoUnit;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class LitexlMapperIntegrationTest {

    @TempDir
    Path tempDir;

    // Custom converter - must be public for reflection access
    public static class MoneyConverter implements LitexlConverter<Money> {
        @Override
        public Money fromCell(CellValue value) {
            return switch (value) {
                case CellValue.Number n -> new Money(n.value());
                default -> null;
            };
        }

        @Override
        public CellValue toCell(Money value) {
            return value == null ? new CellValue.Empty() : new CellValue.Number(value.amount());
        }
    }

    record Money(double amount) {}

    // LocalDate converter - handles both Date and Number (Excel serial date) cell values
    // Required because XlsxReader doesn't detect date formats and reads dates as numbers
    public static class LocalDateConverter implements LitexlConverter<LocalDate> {
        // Excel's epoch: December 31, 1899 (day 1 = January 1, 1900)
        private static final LocalDate EXCEL_EPOCH = LocalDate.of(1899, 12, 31);
        private static final int EXCEL_FAKE_LEAP_DAY = 60;

        @Override
        public LocalDate fromCell(CellValue value) {
            return switch (value) {
                case CellValue.Date d -> d.value().toLocalDate();
                case CellValue.Number n -> fromExcelDate(n.value());
                default -> null;
            };
        }

        @Override
        public CellValue toCell(LocalDate value) {
            if (value == null) {
                return new CellValue.Empty();
            }
            return new CellValue.Date(value.atStartOfDay());
        }

        private LocalDate fromExcelDate(double excelDate) {
            long days = (long) excelDate;
            if (days >= EXCEL_FAKE_LEAP_DAY) {
                days--;
            }
            return EXCEL_EPOCH.plusDays(days);
        }
    }

    // LocalDateTime converter - handles both Date and Number (Excel serial date) cell values
    public static class LocalDateTimeConverter implements LitexlConverter<LocalDateTime> {
        private static final LocalDate EXCEL_EPOCH = LocalDate.of(1899, 12, 31);
        private static final int EXCEL_FAKE_LEAP_DAY = 60;
        private static final double SECONDS_PER_DAY = 86400.0;

        @Override
        public LocalDateTime fromCell(CellValue value) {
            return switch (value) {
                case CellValue.Date d -> d.value();
                case CellValue.Number n -> fromExcelDate(n.value());
                default -> null;
            };
        }

        @Override
        public CellValue toCell(LocalDateTime value) {
            if (value == null) {
                return new CellValue.Empty();
            }
            return new CellValue.Date(value);
        }

        private LocalDateTime fromExcelDate(double excelDate) {
            long days = (long) excelDate;
            double timeFraction = excelDate - days;

            if (days >= EXCEL_FAKE_LEAP_DAY) {
                days--;
            }

            LocalDate date = EXCEL_EPOCH.plusDays(days);

            int totalSeconds = (int) Math.round(timeFraction * SECONDS_PER_DAY);
            int hours = totalSeconds / 3600;
            int minutes = (totalSeconds % 3600) / 60;
            int seconds = totalSeconds % 60;

            LocalTime time = LocalTime.of(hours, minutes, seconds);

            return LocalDateTime.of(date, time);
        }
    }

    // Style provider - must be public for reflection access
    public static class HeaderStyle implements LitexlStyleProvider {
        @Override
        public Style provide() {
            return Style.builder().bold(true).build();
        }
    }

    @LitexlRow
    record Employee(
        @LitexlColumn(index = 0, header = "Name")
        @LitexlStyle(HeaderStyle.class)
        String name,

        @LitexlColumn(index = 1, header = "Salary", converter = MoneyConverter.class)
        Money salary,

        @LitexlColumn(index = 2, header = "HireDate", converter = LocalDateConverter.class)
        LocalDate hireDate,

        @LitexlColumn(index = 3, header = "Active")
        boolean active
    ) {}

    // Summary uses @LitexlRow with @LitexlColumn for row-based mapping
    @LitexlRow
    record Summary(
        @LitexlColumn(index = 0, header = "Title")
        String title,

        @LitexlColumn(index = 1, header = "TotalEmployees")
        int totalEmployees,

        @LitexlColumn(index = 2, header = "GeneratedAt", converter = LocalDateTimeConverter.class)
        LocalDateTime generatedAt
    ) {}

    @LitexlWorkbook
    record CompanyReport(
        @LitexlSheet(name = "Employees")
        List<Employee> employees,

        @LitexlSheet(name = "Summary")
        Summary summary
    ) {}

    // Records for offsetSupport test - must be at class level for reflection to work
    @LitexlRow
    record Item(
        @LitexlColumn(index = 0, header = "Product") String product,
        @LitexlColumn(index = 1, header = "Price") double price
    ) {}

    // Note: dataStartColumn offset is used for writing but reading doesn't account for it yet,
    // so we use row offset only (headerRow/dataStartRow) for the round-trip test
    @LitexlWorkbook
    record OffsetReport(
        @LitexlSheet(name = "Items", headerRow = 2, dataStartRow = 3)
        List<Item> items
    ) {}

    @Test
    void fullRoundTrip() throws Exception {
        var now = LocalDateTime.now().withNano(0);
        var report = new CompanyReport(
            List.of(
                new Employee("Alice", new Money(75000), LocalDate.of(2020, 1, 15), true),
                new Employee("Bob", new Money(85000), LocalDate.of(2019, 6, 1), true),
                new Employee("Charlie", new Money(65000), LocalDate.of(2021, 3, 10), false)
            ),
            new Summary("Q4 Report", 3, now)
        );

        var path = tempDir.resolve("company-report.xlsx");

        // Write
        LitexlMapper.write(report, path);
        assertTrue(path.toFile().exists());

        // Read back
        var result = LitexlMapper.read(path, CompanyReport.class);

        // Verify employees
        assertEquals(3, result.employees().size());

        var alice = result.employees().get(0);
        assertEquals("Alice", alice.name());
        assertEquals(75000, alice.salary().amount());
        assertEquals(LocalDate.of(2020, 1, 15), alice.hireDate());
        assertTrue(alice.active());

        var charlie = result.employees().get(2);
        assertEquals("Charlie", charlie.name());
        assertFalse(charlie.active());

        // Verify summary
        assertEquals("Q4 Report", result.summary().title());
        assertEquals(3, result.summary().totalEmployees());
        assertEquals(now, result.summary().generatedAt());
    }

    @Test
    void offsetSupport() throws Exception {
        var report = new OffsetReport(List.of(
            new Item("Widget", 9.99),
            new Item("Gadget", 19.99)
        ));

        var path = tempDir.resolve("offset-report.xlsx");
        LitexlMapper.write(report, path);

        var result = LitexlMapper.read(path, OffsetReport.class);
        assertEquals(2, result.items().size());
        assertEquals("Widget", result.items().get(0).product());
        assertEquals(19.99, result.items().get(1).price());
    }

    @Test
    void emptyWorkbook() throws Exception {
        var report = new CompanyReport(List.of(), new Summary("Empty", 0, LocalDateTime.now().withNano(0)));
        var path = tempDir.resolve("empty.xlsx");

        LitexlMapper.write(report, path);
        var result = LitexlMapper.read(path, CompanyReport.class);

        assertTrue(result.employees().isEmpty());
        assertEquals("Empty", result.summary().title());
    }
}
