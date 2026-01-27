package com.beingidly.litexl.examples.mapper;

import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.mapper.LitexlColumn;
import com.beingidly.litexl.mapper.LitexlMapper;
import com.beingidly.litexl.mapper.LitexlRow;
import com.beingidly.litexl.mapper.LitexlSheet;
import com.beingidly.litexl.mapper.LitexlStyle;
import com.beingidly.litexl.mapper.LitexlStyleProvider;
import com.beingidly.litexl.mapper.LitexlWorkbook;
import com.beingidly.litexl.style.BorderStyle;
import com.beingidly.litexl.style.HAlign;
import com.beingidly.litexl.style.Style;
import com.beingidly.litexl.style.VAlign;

import java.nio.file.Path;
import java.util.List;

/**
 * Example: Applying styles using LitexlStyleProvider.
 *
 * This example demonstrates:
 * - Creating custom style providers
 * - Applying styles to columns via @LitexlStyle annotation
 */
public class Ex03_StyleProvider {

    // Header style provider - bold, blue background, white text
    public static class HeaderStyle implements LitexlStyleProvider {
        @Override
        public Style provide() {
            return Style.builder()
                .bold(true)
                .fill(0xFF4472C4)      // Excel blue
                .color(0xFFFFFFFF)     // White text
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .border(BorderStyle.THIN, 0xFF000000)
                .build();
        }
    }

    // Currency style provider - number format with dollar sign
    public static class CurrencyStyle implements LitexlStyleProvider {
        @Override
        public Style provide() {
            return Style.builder()
                .format("$#,##0.00")
                .align(HAlign.RIGHT, VAlign.MIDDLE)
                .build();
        }
    }

    // Row record with styled columns
    @LitexlRow
    public record SalesRecord(
        @LitexlColumn(index = 0, header = "Product")
        @LitexlStyle(HeaderStyle.class)
        String product,

        @LitexlColumn(index = 1, header = "Quantity")
        @LitexlStyle(HeaderStyle.class)
        int quantity,

        @LitexlColumn(index = 2, header = "Revenue")
        @LitexlStyle(CurrencyStyle.class)
        double revenue
    ) {}

    @LitexlWorkbook
    public record SalesReport(
        @LitexlSheet(name = "Sales")
        List<SalesRecord> sales
    ) {}

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex03_style_provider.xlsx");

        // Create sample data
        List<SalesRecord> sales = List.of(
            new SalesRecord("Laptop", 10, 9999.90),
            new SalesRecord("Mouse", 50, 1499.50),
            new SalesRecord("Keyboard", 25, 1987.25)
        );

        SalesReport report = new SalesReport(sales);

        // Write to Excel - styles will be applied automatically
        LitexlMapper.write(report, outputPath);
        ExampleUtils.printCreated(outputPath);

        System.out.println();
        System.out.println("Open the file to see styled headers and currency format.");
    }
}
