# litexl

Pure Java library for reading and writing Excel XLSX files. Lightweight, fast, with no external dependencies beyond the Java standard library.

## Features

- **Read/Write XLSX files** - Full support for reading and writing Excel 2007+ format
- **Object Mapping** - Annotation-based mapping between Java objects and Excel sheets
- **Styling** - Fonts, borders, fills, number formats, alignment
- **Formulas** - Write formula expressions to cells
- **Cell merging** - Merge ranges of cells
- **Data validation** - List, number range, text length, custom formula validations
- **Conditional formatting** - Apply conditional styles based on cell values
- **AutoFilter** - Add filter dropdowns to data ranges
- **Document encryption** - AES-128/256 encryption with password protection (ECMA-376 Agile)
- **Sheet protection** - Protect sheets with password and permission options
- **Type-safe API** - Sealed interfaces for cell values with pattern matching
- **GraalVM Native Image compatible**

## Requirements

- Java 21+
- Gradle 8.0+ (for building)

## Installation

### Gradle

```kotlin
dependencies {
    implementation("com.beingidly:litexl:0.1.0")
}
```

### Maven

```xml
<dependency>
    <groupId>com.beingidly</groupId>
    <artifactId>litexl</artifactId>
    <version>0.1.0</version>
</dependency>
```

## Quick Start

### Creating a Workbook

```java
import com.beingidly.litexl.*;
import java.nio.file.Path;

try (Workbook wb = Workbook.create()) {
    Sheet sheet = wb.addSheet("My Sheet");

    // Set cell values
    sheet.cell(0, 0).set("Name");
    sheet.cell(0, 1).set("Age");
    sheet.cell(0, 2).set("Active");

    sheet.cell(1, 0).set("Alice");
    sheet.cell(1, 1).set(30);
    sheet.cell(1, 2).set(true);

    sheet.cell(2, 0).set("Bob");
    sheet.cell(2, 1).set(25);
    sheet.cell(2, 2).set(false);

    // Save to file
    wb.save(Path.of("output.xlsx"));
}
```

### Reading a Workbook

```java
try (Workbook wb = Workbook.open(Path.of("input.xlsx"))) {
    Sheet sheet = wb.getSheet(0);

    for (var row : sheet.rows().values()) {
        for (var cell : row.cells().values()) {
            System.out.println(cell.type() + ": " + cell.rawValue());
        }
    }
}
```

## Object Mapping with LitexlMapper

LitexlMapper provides annotation-based object mapping for Excel files. Define Java records with annotations and let litexl handle the conversion.

### Basic Mapping

```java
import com.beingidly.litexl.mapper.*;
import java.util.List;

// Define a row record
@LitexlRow
public record Person(
    @LitexlColumn(index = 0, header = "Name")
    String name,

    @LitexlColumn(index = 1, header = "Age")
    int age,

    @LitexlColumn(index = 2, header = "Email")
    String email
) {}

// Define a workbook record
@LitexlWorkbook
public record PeopleReport(
    @LitexlSheet(name = "People")
    List<Person> people
) {}

// Write to Excel
List<Person> people = List.of(
    new Person("Alice", 30, "alice@example.com"),
    new Person("Bob", 25, "bob@example.com")
);
LitexlMapper.write(new PeopleReport(people), Path.of("people.xlsx"));

// Read from Excel
PeopleReport report = LitexlMapper.read(Path.of("people.xlsx"), PeopleReport.class);
```

### Custom Type Converters

Convert between Excel values and domain types:

```java
// Custom domain type
public record Money(String currency, double amount) {}

// Custom converter
public class MoneyConverter implements LitexlConverter<Money> {
    @Override
    public Money fromCell(CellValue value) {
        return switch (value) {
            case CellValue.Text(String text) -> {
                String[] parts = text.split(" ", 2);
                yield new Money(parts[0], Double.parseDouble(parts[1]));
            }
            case CellValue.Number(double num) -> new Money("USD", num);
            default -> null;
        };
    }

    @Override
    public CellValue toCell(Money value) {
        return value != null
            ? new CellValue.Text(value.currency() + " " + value.amount())
            : new CellValue.Empty();
    }
}

// Use in row record
@LitexlRow
public record Product(
    @LitexlColumn(index = 0, header = "Name")
    String name,

    @LitexlColumn(index = 1, header = "Price", converter = MoneyConverter.class)
    Money price
) {}
```

### Style Providers

Apply styles using annotation-based providers:

```java
public class HeaderStyle implements LitexlStyleProvider {
    @Override
    public Style provide() {
        return Style.builder()
            .bold(true)
            .fill(0xFF4472C4)
            .color(0xFFFFFFFF)
            .align(HAlign.CENTER, VAlign.MIDDLE)
            .build();
    }
}

@LitexlRow
public record SalesRecord(
    @LitexlColumn(index = 0, header = "Product")
    @LitexlStyle(HeaderStyle.class)
    String product,

    @LitexlColumn(index = 1, header = "Revenue")
    double revenue
) {}
```

### Auto Region Detection

Read data that doesn't start at A1:

```java
@LitexlWorkbook
public record Report(
    @LitexlSheet(name = "Data", regionDetection = RegionDetection.AUTO)
    List<Row> rows
) {}
```

### Mapper Configuration

```java
LitexlMapper mapper = LitexlMapper.builder()
    .dateFormat("yyyy-MM-dd")
    .nullStrategy(NullStrategy.EMPTY_CELL)
    .build();

Report report = mapper.readFile(path, Report.class);
mapper.writeFile(report, outputPath);
```

## Styling Cells

```java
import com.beingidly.litexl.style.*;

try (Workbook wb = Workbook.create()) {
    Sheet sheet = wb.addSheet("Styled");

    // Create styles
    int headerStyle = wb.addStyle(Style.builder()
        .bold(true)
        .fill(0xFF4472C4)  // Blue background
        .color(0xFFFFFFFF)  // White text
        .align(HAlign.CENTER, VAlign.MIDDLE)
        .build());

    int numberStyle = wb.addStyle(Style.builder()
        .format("#,##0.00")
        .align(HAlign.RIGHT, VAlign.BOTTOM)
        .border(BorderStyle.THIN, 0xFF000000)
        .build());

    // Apply styles
    sheet.cell(0, 0).set("Revenue").style(headerStyle);
    sheet.cell(1, 0).set(1234567.89).style(numberStyle);

    wb.save(Path.of("styled.xlsx"));
}
```

## Formulas

```java
sheet.cell(0, 0).set(10);
sheet.cell(0, 1).set(20);
sheet.cell(0, 2).setFormula("A1+B1");
sheet.cell(1, 0).setFormula("SUM(A1:B1)");
```

## Merging Cells

```java
sheet.cell(0, 0).set("Merged Header");
sheet.merge(0, 0, 0, 5);  // Merge A1:F1
```

## Data Validation

```java
import com.beingidly.litexl.format.DataValidation;

// Dropdown list
sheet.addValidation(DataValidation.list(
    CellRange.of(1, 0, 100, 0),  // A2:A101
    "Yes", "No", "Maybe"
));

// Number range
sheet.addValidation(DataValidation.wholeNumber(
    CellRange.of(1, 1, 100, 1),  // B2:B101
    1, 100
));
```

## AutoFilter

```java
sheet.setAutoFilter(0, 0, 100, 5);  // A1:F101
```

## Encryption

```java
import com.beingidly.litexl.crypto.EncryptionOptions;

// Save with encryption
wb.save(Path.of("encrypted.xlsx"),
    EncryptionOptions.aes256("password123"));

// Open encrypted file
try (Workbook wb = Workbook.open(
        Path.of("encrypted.xlsx"), "password123")) {
    // ...
}
```

## Sheet Protection

Protect sheets with passwords and configurable permissions:

```java
import com.beingidly.litexl.crypto.SheetProtection;

// Protect without password
sheet.protect(SheetProtection.defaults());

// Protect with password and custom permissions
sheet.protectionManager().protect("password".toCharArray(),
    SheetProtection.builder()
        .formatCells(true)
        .insertRows(true)
        .deleteRows(true)
        .sort(true)
        .autoFilter(true)
        .build());

// Unprotect
sheet.protectionManager().unprotect("password".toCharArray());
```

## Type-Safe Cell Values

litexl uses sealed interfaces for type-safe cell values:

```java
CellValue value = cell.value();

String result = switch (value) {
    case CellValue.Empty _ -> "empty";
    case CellValue.Text t -> "text: " + t.value();
    case CellValue.Number n -> "number: " + n.value();
    case CellValue.Bool b -> "bool: " + b.value();
    case CellValue.Date d -> "date: " + d.value();
    case CellValue.Formula f -> "formula: " + f.expression();
    case CellValue.Error e -> "error: " + e.code();
};
```

## Building from Source

```bash
git clone https://github.com/beingidly/litexl.git
cd litexl
./gradlew build
```

### Running Tests

```bash
./gradlew test
```

### Running Benchmarks

```bash
./gradlew jmh
```

### Building Native Image

```bash
./gradlew nativeCompile
```

## Architecture

```
litexl
├── com.beingidly.litexl              # Public API
│   ├── Workbook                      # Main entry point
│   ├── Sheet                         # Worksheet
│   ├── Row                           # Row container
│   ├── Cell                          # Cell with value and style
│   ├── CellValue                     # Sealed interface for values
│   ├── CellRange                     # Range reference
│   └── CellType                      # Enum of cell types
├── com.beingidly.litexl.mapper       # Object Mapping
│   ├── LitexlMapper                  # Main mapper class
│   ├── @LitexlWorkbook               # Workbook annotation
│   ├── @LitexlSheet                  # Sheet annotation
│   ├── @LitexlRow                    # Row annotation
│   ├── @LitexlColumn                 # Column annotation
│   ├── @LitexlStyle                  # Style annotation
│   ├── LitexlConverter               # Custom type converter
│   └── LitexlStyleProvider           # Custom style provider
├── com.beingidly.litexl.style        # Styling
│   ├── Style                         # Cell style
│   ├── Font                          # Font properties
│   ├── Border                        # Border properties
│   └── Alignment                     # Alignment settings
├── com.beingidly.litexl.format       # Formatting
│   ├── ConditionalFormat             # Conditional formatting
│   ├── DataValidation                # Data validation rules
│   └── AutoFilter                    # AutoFilter settings
├── com.beingidly.litexl.crypto       # Encryption
│   ├── EncryptionOptions             # Encryption settings
│   └── SheetProtection               # Sheet protection settings
└── com.beingidly.litexl.internal     # Internal implementation
    ├── xlsx                          # XLSX read/write
    ├── xml                           # XML processing
    ├── zip                           # ZIP handling
    └── crypto                        # Crypto utilities
```

## License

Apache License 2.0
