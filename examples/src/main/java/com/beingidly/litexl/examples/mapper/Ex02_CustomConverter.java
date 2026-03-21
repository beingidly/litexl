package com.beingidly.litexl.examples.mapper;

import com.beingidly.litexl.CellValue;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.mapper.LitexlColumn;
import com.beingidly.litexl.mapper.LitexlConverter;
import com.beingidly.litexl.mapper.LitexlMapper;
import com.beingidly.litexl.mapper.LitexlRow;
import com.beingidly.litexl.mapper.LitexlSheet;
import com.beingidly.litexl.mapper.LitexlWorkbook;

import java.nio.file.Path;
import java.util.List;

/**
 * Example: Custom type converters for domain types.
 *
 * This example demonstrates:
 * - Creating custom converters for domain types
 * - Converting between Excel values and Java objects
 */
public class Ex02_CustomConverter {

    /** Example runner. */
    private Ex02_CustomConverter() {}

    /**
     * A money value with currency and amount.
     *
     * @param currency the currency code (e.g. "USD", "EUR")
     * @param amount   the monetary amount
     */
    public record Money(String currency, double amount) {
        @Override
        public String toString() {
            return currency + " " + String.format("%.2f", amount);
        }
    }

    /** Custom converter for Money type. */
    public static class MoneyConverter implements LitexlConverter<Money> {

        /** Creates a new {@code MoneyConverter}. */
        public MoneyConverter() {}

        @Override
        public Money fromCell(CellValue value) {
            // Parse from string format: "USD 100.00" or just number
            return switch (value) {
                case CellValue.Text(String text) -> {
                    String[] parts = text.split(" ", 2);
                    if (parts.length == 2) {
                        yield new Money(parts[0], Double.parseDouble(parts[1]));
                    }
                    yield new Money("USD", Double.parseDouble(text));
                }
                case CellValue.Number(double num) -> new Money("USD", num);
                default -> null;
            };
        }

        @Override
        public CellValue toCell(Money value) {
            if (value == null) {
                return new CellValue.Empty();
            }
            // Store as string format: "USD 100.00"
            return new CellValue.Text(value.toString());
        }
    }

    /**
     * A product row record using custom converter.
     *
     * @param name  the product name
     * @param price the product price as a {@link Money} value
     * @param stock the number of items in stock
     */
    @LitexlRow
    public record Product(
        @LitexlColumn(index = 0, header = "Product Name")
        String name,

        @LitexlColumn(index = 1, header = "Price", converter = MoneyConverter.class)
        Money price,

        @LitexlColumn(index = 2, header = "Stock")
        int stock
    ) {}

    /**
     * A workbook record containing products.
     *
     * @param products the list of products in the catalog
     */
    @LitexlWorkbook
    public record ProductCatalog(
        @LitexlSheet(name = "Products")
        List<Product> products
    ) {}

    /**
     * Runs the example.
     *
     * @param args command-line arguments (not used)
     */
    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex02_custom_converter.xlsx");

        // Create sample data with custom Money type
        List<Product> products = List.of(
            new Product("Laptop", new Money("USD", 999.99), 50),
            new Product("Mouse", new Money("USD", 29.99), 200),
            new Product("Keyboard", new Money("EUR", 79.50), 75)
        );

        ProductCatalog catalog = new ProductCatalog(products);

        // Write to Excel
        LitexlMapper.write(catalog, outputPath);
        ExampleUtils.printCreated(outputPath);

        // Read back
        System.out.println();
        System.out.println("Reading back with custom converter:");

        ProductCatalog readBack = LitexlMapper.read(outputPath, ProductCatalog.class);

        for (Product product : readBack.products()) {
            System.out.println("  " + product.name() + " - " + product.price() + " (stock: " + product.stock() + ")");
        }
    }
}
