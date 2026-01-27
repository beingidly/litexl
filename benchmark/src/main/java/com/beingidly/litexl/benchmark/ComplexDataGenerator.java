package com.beingidly.litexl.benchmark;

import com.beingidly.litexl.CellRange;
import com.beingidly.litexl.CellValue;
import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.crypto.SheetProtection;
import com.beingidly.litexl.format.ConditionalFormat;
import com.beingidly.litexl.format.DataValidation;
import com.beingidly.litexl.style.BorderStyle;
import com.beingidly.litexl.style.HAlign;
import com.beingidly.litexl.style.Style;
import com.beingidly.litexl.style.VAlign;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Random;

/**
 * Generates complex workbook data for comprehensive benchmarking.
 *
 * <p>Sheet configuration (5 sheets, ~9,300 rows total):
 * <ul>
 *   <li>Summary (100 rows): Dashboard with merged cells, formulas, conditional formatting</li>
 *   <li>Sales (3,000 rows, 15 cols): Orders with autofilter, date/currency/percent formatting</li>
 *   <li>Customers (3,000 rows, 12 cols): Customer data with validation dropdown, borders</li>
 *   <li>Products (3,000 rows, 10 cols): Products with category colors, column widths</li>
 *   <li>Config (100 rows): Key-Value-Description with booleans, errors, sheet protection</li>
 * </ul>
 */
public class ComplexDataGenerator {

    private static final int SUMMARY_ROWS = 100;
    private static final int SALES_ROWS = 3000;
    private static final int CUSTOMERS_ROWS = 3000;
    private static final int PRODUCTS_ROWS = 3000;
    private static final int CONFIG_ROWS = 100;

    private static final int SALES_COLS = 15;
    private static final int CUSTOMERS_COLS = 12;
    private static final int PRODUCTS_COLS = 10;
    private static final int CONFIG_COLS = 3;
    private static final int SUMMARY_COLS = 5;

    private static final String[] REGIONS = {"North", "South", "East", "West", "Central"};
    private static final String[] STATUSES = {"Pending", "Shipped", "Delivered", "Cancelled"};
    private static final String[] PAYMENTS = {"Credit Card", "PayPal", "Bank Transfer", "Cash"};
    private static final String[] GRADES = {"A", "B", "C"};
    private static final String[] COUNTRIES = {"USA", "Canada", "UK", "Germany", "France", "Japan", "Australia"};
    private static final String[] CATEGORIES = {"Electronics", "Clothing", "Food", "Books", "Home"};

    // Package-private for ComplexCompare to reset
    Random random = new Random(42);

    /**
     * Generates a complex workbook using litexl API.
     */
    public void generateLitexl(Path path) throws Exception {
        try (Workbook wb = Workbook.create()) {
            // Create styles
            int headerStyle = wb.addStyle(Style.builder()
                .bold(true)
                .fill(0xFF4472C4)
                .color(0xFFFFFFFF)
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            int currencyStyle = wb.addStyle(Style.builder()
                .format("$#,##0.00")
                .align(HAlign.RIGHT, VAlign.MIDDLE)
                .build());

            int percentStyle = wb.addStyle(Style.builder()
                .format("0.00%")
                .align(HAlign.RIGHT, VAlign.MIDDLE)
                .build());

            int dateStyle = wb.addStyle(Style.builder()
                .format("yyyy-mm-dd")
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .build());

            int borderStyle = wb.addStyle(Style.builder()
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            int highlightStyle = wb.addStyle(Style.builder()
                .fill(0xFFFFC000)
                .bold(true)
                .build());

            // Category color styles
            int[] categoryStyles = new int[CATEGORIES.length];
            int[] categoryColors = {0xFFE2EFDA, 0xFFFCE4D6, 0xFFDDEBF7, 0xFFFFF2CC, 0xFFE4DFEC};
            for (int i = 0; i < CATEGORIES.length; i++) {
                categoryStyles[i] = wb.addStyle(Style.builder()
                    .fill(categoryColors[i])
                    .build());
            }

            // Generate sheets
            generateSummaryLitexl(wb, headerStyle, currencyStyle, highlightStyle);
            generateSalesLitexl(wb, headerStyle, currencyStyle, percentStyle, dateStyle);
            generateCustomersLitexl(wb, headerStyle, dateStyle, borderStyle);
            generateProductsLitexl(wb, headerStyle, currencyStyle, percentStyle, categoryStyles);
            generateConfigLitexl(wb, headerStyle, borderStyle);

            wb.save(path);
        }
    }

    private void generateSummaryLitexl(Workbook wb, int headerStyle, int currencyStyle, int highlightStyle) {
        Sheet sheet = wb.addSheet("Summary");

        // Header with merged cells
        sheet.cell(0, 0).set("Dashboard Summary").style(headerStyle);
        sheet.merge(0, 0, 0, 4);

        // Section headers
        String[] sections = {"Sales Overview", "Customer Metrics", "Product Stats", "Regional Summary"};
        int row = 2;
        for (String section : sections) {
            sheet.cell(row, 0).set(section).style(headerStyle);
            sheet.merge(row, 0, row, 2);

            // Metrics with formulas referencing other sheets
            sheet.cell(row + 1, 0).set("Total Count");
            sheet.cell(row + 1, 1).setFormula("COUNTA(Sales!A:A)-1");

            sheet.cell(row + 2, 0).set("Total Value");
            sheet.cell(row + 2, 1).setFormula("SUM(Sales!H:H)").style(currencyStyle);

            sheet.cell(row + 3, 0).set("Average");
            sheet.cell(row + 3, 1).setFormula("AVERAGE(Sales!H:H)").style(currencyStyle);

            row += 5;
        }

        // Fill remaining rows with metrics
        for (int r = row; r < SUMMARY_ROWS; r++) {
            sheet.cell(r, 0).set("Metric " + (r - row + 1));
            sheet.cell(r, 1).set(random.nextDouble() * 10000);
            sheet.cell(r, 2).set(random.nextDouble() * 100);
            sheet.cell(r, 3).setFormula("B" + (r + 1) + "*C" + (r + 1));
            sheet.cell(r, 4).set(random.nextBoolean() ? "Pass" : "Fail");
        }

        // Conditional formatting for high values
        CellRange valueRange = CellRange.of(2, 1, SUMMARY_ROWS - 1, 1);
        sheet.addConditionalFormat(ConditionalFormat.greaterThan(valueRange, 5000, highlightStyle));

        // Column widths
        sheet.setColumnWidth(0, 20.0);
        sheet.setColumnWidth(1, 15.0);
        sheet.setColumnWidth(2, 15.0);
        sheet.setColumnWidth(3, 15.0);
        sheet.setColumnWidth(4, 10.0);
    }

    private void generateSalesLitexl(Workbook wb, int headerStyle, int currencyStyle,
                                      int percentStyle, int dateStyle) {
        Sheet sheet = wb.addSheet("Sales");

        // Headers
        String[] headers = {"OrderID", "Date", "Customer", "Product", "Quantity", "UnitPrice",
            "Discount", "Total", "Tax", "NetAmount", "Region", "Status", "ShipDate", "Payment", "Notes"};
        for (int c = 0; c < headers.length; c++) {
            sheet.cell(0, c).set(headers[c]).style(headerStyle);
        }

        // Data rows
        LocalDate baseDate = LocalDate.of(2024, 1, 1);
        for (int r = 1; r <= SALES_ROWS; r++) {
            int row = r;
            sheet.cell(row, 0).set("ORD-" + String.format("%06d", r));
            sheet.cell(row, 1).set(baseDate.plusDays(random.nextInt(365)).atStartOfDay()).style(dateStyle);
            sheet.cell(row, 2).set("Customer " + (random.nextInt(1000) + 1));
            sheet.cell(row, 3).set("Product " + (random.nextInt(500) + 1));
            sheet.cell(row, 4).set(random.nextInt(100) + 1);
            sheet.cell(row, 5).set(Math.round(random.nextDouble() * 500 * 100) / 100.0).style(currencyStyle);
            sheet.cell(row, 6).set(Math.round(random.nextDouble() * 0.3 * 100) / 100.0).style(percentStyle);

            // Formula columns
            int excelRow = row + 1;
            sheet.cell(row, 7).setFormula("E" + excelRow + "*F" + excelRow + "*(1-G" + excelRow + ")").style(currencyStyle);
            sheet.cell(row, 8).setFormula("H" + excelRow + "*0.08").style(currencyStyle);
            sheet.cell(row, 9).setFormula("H" + excelRow + "+I" + excelRow).style(currencyStyle);

            sheet.cell(row, 10).set(REGIONS[random.nextInt(REGIONS.length)]);
            sheet.cell(row, 11).set(STATUSES[random.nextInt(STATUSES.length)]);
            sheet.cell(row, 12).set(baseDate.plusDays(random.nextInt(365) + 1).atStartOfDay()).style(dateStyle);
            sheet.cell(row, 13).set(PAYMENTS[random.nextInt(PAYMENTS.length)]);
            sheet.cell(row, 14).set("Note for order " + r);
        }

        // AutoFilter
        sheet.setAutoFilter(0, 0, SALES_ROWS, SALES_COLS - 1);

        // Column widths
        double[] widths = {12, 12, 15, 15, 10, 12, 10, 12, 10, 12, 10, 10, 12, 15, 20};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, widths[c]);
        }
    }

    private void generateCustomersLitexl(Workbook wb, int headerStyle, int dateStyle, int borderStyle) {
        Sheet sheet = wb.addSheet("Customers");

        // Headers
        String[] headers = {"ID", "Name", "Email", "Phone", "Address", "City",
            "Country", "JoinDate", "Grade", "CreditLimit", "IsActive", "Memo"};
        for (int c = 0; c < headers.length; c++) {
            sheet.cell(0, c).set(headers[c]).style(headerStyle);
        }

        // Data rows
        LocalDate baseDate = LocalDate.of(2020, 1, 1);
        for (int r = 1; r <= CUSTOMERS_ROWS; r++) {
            int row = r;
            sheet.cell(row, 0).set("CUST-" + String.format("%05d", r)).style(borderStyle);
            sheet.cell(row, 1).set("Customer Name " + r).style(borderStyle);
            sheet.cell(row, 2).set("customer" + r + "@example.com").style(borderStyle);
            sheet.cell(row, 3).set("+1-555-" + String.format("%04d", random.nextInt(10000))).style(borderStyle);
            sheet.cell(row, 4).set(random.nextInt(9999) + " Main Street").style(borderStyle);
            sheet.cell(row, 5).set("City " + (random.nextInt(100) + 1)).style(borderStyle);
            sheet.cell(row, 6).set(COUNTRIES[random.nextInt(COUNTRIES.length)]).style(borderStyle);
            sheet.cell(row, 7).set(baseDate.plusDays(random.nextInt(1500)).atStartOfDay()).style(dateStyle);
            sheet.cell(row, 8).set(GRADES[random.nextInt(GRADES.length)]).style(borderStyle);
            sheet.cell(row, 9).set(Math.round(random.nextDouble() * 50000 * 100) / 100.0).style(borderStyle);
            sheet.cell(row, 10).set(random.nextBoolean()).style(borderStyle);
            sheet.cell(row, 11).set("Memo for customer " + r).style(borderStyle);
        }

        // Data validation for Grade column
        CellRange gradeRange = CellRange.of(1, 8, CUSTOMERS_ROWS, 8);
        sheet.addValidation(DataValidation.list(gradeRange, "A", "B", "C"));

        // Column widths
        double[] widths = {12, 20, 25, 15, 20, 12, 12, 12, 8, 12, 10, 25};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, widths[c]);
        }
    }

    private void generateProductsLitexl(Workbook wb, int headerStyle, int currencyStyle,
                                         int percentStyle, int[] categoryStyles) {
        Sheet sheet = wb.addSheet("Products");

        // Headers
        String[] headers = {"SKU", "Name", "Category", "Price", "Cost",
            "Margin", "Stock", "MinStock", "Reorder", "LastUpdate"};
        for (int c = 0; c < headers.length; c++) {
            sheet.cell(0, c).set(headers[c]).style(headerStyle);
        }

        // Data rows
        LocalDateTime baseDate = LocalDateTime.of(2024, 1, 1, 0, 0);
        for (int r = 1; r <= PRODUCTS_ROWS; r++) {
            int row = r;
            int categoryIdx = random.nextInt(CATEGORIES.length);
            int categoryStyle = categoryStyles[categoryIdx];

            sheet.cell(row, 0).set("SKU-" + String.format("%06d", r)).style(categoryStyle);
            sheet.cell(row, 1).set("Product Name " + r).style(categoryStyle);
            sheet.cell(row, 2).set(CATEGORIES[categoryIdx]).style(categoryStyle);

            double price = Math.round(random.nextDouble() * 1000 * 100) / 100.0;
            double cost = Math.round(price * (0.4 + random.nextDouble() * 0.3) * 100) / 100.0;
            sheet.cell(row, 3).set(price).style(currencyStyle);
            sheet.cell(row, 4).set(cost).style(currencyStyle);

            // Margin formula
            int excelRow = row + 1;
            sheet.cell(row, 5).setFormula("(D" + excelRow + "-E" + excelRow + ")/D" + excelRow).style(percentStyle);

            int stock = random.nextInt(1000);
            int minStock = random.nextInt(50) + 10;
            sheet.cell(row, 6).set(stock).style(categoryStyle);
            sheet.cell(row, 7).set(minStock).style(categoryStyle);

            // Reorder formula (IF formula)
            sheet.cell(row, 8).setFormula("IF(G" + excelRow + "<H" + excelRow + ",\"Yes\",\"No\")").style(categoryStyle);
            sheet.cell(row, 9).set(baseDate.plusDays(random.nextInt(365))).style(categoryStyle);
        }

        // Column widths
        double[] widths = {15, 25, 12, 12, 12, 10, 10, 10, 10, 15};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, widths[c]);
        }
    }

    private void generateConfigLitexl(Workbook wb, int headerStyle, int borderStyle) {
        Sheet sheet = wb.addSheet("Config");

        // Headers
        String[] headers = {"Key", "Value", "Description"};
        for (int c = 0; c < headers.length; c++) {
            sheet.cell(0, c).set(headers[c]).style(headerStyle);
        }

        // Configuration data with various types
        int row = 1;

        // Boolean values
        sheet.cell(row, 0).set("EnableFeatureA").style(borderStyle);
        sheet.cell(row, 1).set(true).style(borderStyle);
        sheet.cell(row++, 2).set("Enable feature A").style(borderStyle);

        sheet.cell(row, 0).set("EnableFeatureB").style(borderStyle);
        sheet.cell(row, 1).set(false).style(borderStyle);
        sheet.cell(row++, 2).set("Enable feature B").style(borderStyle);

        // Error values
        sheet.cell(row, 0).set("ErrorNA").style(borderStyle);
        sheet.cell(row, 1).setValue(new CellValue.Error("#N/A")).style(borderStyle);
        sheet.cell(row++, 2).set("Not available error").style(borderStyle);

        sheet.cell(row, 0).set("ErrorREF").style(borderStyle);
        sheet.cell(row, 1).setValue(new CellValue.Error("#REF!")).style(borderStyle);
        sheet.cell(row++, 2).set("Reference error").style(borderStyle);

        sheet.cell(row, 0).set("ErrorVALUE").style(borderStyle);
        sheet.cell(row, 1).setValue(new CellValue.Error("#VALUE!")).style(borderStyle);
        sheet.cell(row++, 2).set("Value error").style(borderStyle);

        // Fill remaining rows
        for (; row < CONFIG_ROWS; row++) {
            String key = "Config_" + (row - 5);
            sheet.cell(row, 0).set(key).style(borderStyle);

            // Alternate between different value types
            switch (row % 4) {
                case 0 -> sheet.cell(row, 1).set("StringValue" + row).style(borderStyle);
                case 1 -> sheet.cell(row, 1).set(random.nextDouble() * 1000).style(borderStyle);
                case 2 -> sheet.cell(row, 1).set(random.nextBoolean()).style(borderStyle);
                case 3 -> sheet.cell(row, 1).set(random.nextInt(1000)).style(borderStyle);
            }
            sheet.cell(row, 2).set("Description for " + key).style(borderStyle);
        }

        // Sheet protection
        sheet.protect(SheetProtection.defaults());

        // Column widths
        sheet.setColumnWidth(0, 20.0);
        sheet.setColumnWidth(1, 20.0);
        sheet.setColumnWidth(2, 40.0);
    }

    /**
     * Generates a complex workbook using Apache POI API.
     */
    public void generatePoi(Path path) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            // Create styles
            XSSFCellStyle headerStyle = wb.createCellStyle();
            XSSFFont headerFont = wb.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 68, (byte) 114, (byte) 196}, null));
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            headerStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            headerStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            headerStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);

            XSSFCellStyle currencyStyle = wb.createCellStyle();
            currencyStyle.setDataFormat(wb.createDataFormat().getFormat("$#,##0.00"));
            currencyStyle.setAlignment(HorizontalAlignment.RIGHT);

            XSSFCellStyle percentStyle = wb.createCellStyle();
            percentStyle.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
            percentStyle.setAlignment(HorizontalAlignment.RIGHT);

            XSSFCellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd"));
            dateStyle.setAlignment(HorizontalAlignment.CENTER);

            XSSFCellStyle borderStyle = wb.createCellStyle();
            borderStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            borderStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            borderStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            borderStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);

            XSSFCellStyle highlightStyle = wb.createCellStyle();
            highlightStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 255, (byte) 192, (byte) 0}, null));
            highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            XSSFFont boldFont = wb.createFont();
            boldFont.setBold(true);
            highlightStyle.setFont(boldFont);

            // Category color styles
            XSSFCellStyle[] categoryStyles = new XSSFCellStyle[CATEGORIES.length];
            byte[][] categoryColors = {
                {(byte) 226, (byte) 239, (byte) 218},
                {(byte) 252, (byte) 228, (byte) 214},
                {(byte) 221, (byte) 235, (byte) 247},
                {(byte) 255, (byte) 242, (byte) 204},
                {(byte) 228, (byte) 223, (byte) 236}
            };
            for (int i = 0; i < CATEGORIES.length; i++) {
                categoryStyles[i] = wb.createCellStyle();
                categoryStyles[i].setFillForegroundColor(new XSSFColor(categoryColors[i], null));
                categoryStyles[i].setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            // Generate sheets
            generateSummaryPoi(wb, headerStyle, currencyStyle, highlightStyle);
            generateSalesPoi(wb, headerStyle, currencyStyle, percentStyle, dateStyle);
            generateCustomersPoi(wb, headerStyle, dateStyle, borderStyle);
            generateProductsPoi(wb, headerStyle, currencyStyle, percentStyle, categoryStyles);
            generateConfigPoi(wb, headerStyle, borderStyle);

            try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
                wb.write(fos);
            }
        }
    }

    private void generateSummaryPoi(XSSFWorkbook wb, XSSFCellStyle headerStyle,
                                     XSSFCellStyle currencyStyle, XSSFCellStyle highlightStyle) {
        XSSFSheet sheet = wb.createSheet("Summary");

        // Header with merged cells
        XSSFRow row0 = sheet.createRow(0);
        XSSFCell titleCell = row0.createCell(0);
        titleCell.setCellValue("Dashboard Summary");
        titleCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

        // Section headers
        String[] sections = {"Sales Overview", "Customer Metrics", "Product Stats", "Regional Summary"};
        int rowNum = 2;
        for (String section : sections) {
            XSSFRow sectionRow = sheet.createRow(rowNum);
            XSSFCell sectionCell = sectionRow.createCell(0);
            sectionCell.setCellValue(section);
            sectionCell.setCellStyle(headerStyle);
            sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, 2));

            XSSFRow row1 = sheet.createRow(rowNum + 1);
            row1.createCell(0).setCellValue("Total Count");
            row1.createCell(1).setCellFormula("COUNTA(Sales!A:A)-1");

            XSSFRow row2 = sheet.createRow(rowNum + 2);
            row2.createCell(0).setCellValue("Total Value");
            XSSFCell valueCell = row2.createCell(1);
            valueCell.setCellFormula("SUM(Sales!H:H)");
            valueCell.setCellStyle(currencyStyle);

            XSSFRow row3 = sheet.createRow(rowNum + 3);
            row3.createCell(0).setCellValue("Average");
            XSSFCell avgCell = row3.createCell(1);
            avgCell.setCellFormula("AVERAGE(Sales!H:H)");
            avgCell.setCellStyle(currencyStyle);

            rowNum += 5;
        }

        // Fill remaining rows with metrics
        for (int r = rowNum; r < SUMMARY_ROWS; r++) {
            XSSFRow row = sheet.createRow(r);
            row.createCell(0).setCellValue("Metric " + (r - rowNum + 1));
            row.createCell(1).setCellValue(random.nextDouble() * 10000);
            row.createCell(2).setCellValue(random.nextDouble() * 100);
            row.createCell(3).setCellFormula("B" + (r + 1) + "*C" + (r + 1));
            row.createCell(4).setCellValue(random.nextBoolean() ? "Pass" : "Fail");
        }

        // Conditional formatting
        XSSFSheetConditionalFormatting cf = sheet.getSheetConditionalFormatting();
        XSSFConditionalFormattingRule rule = cf.createConditionalFormattingRule(
            ComparisonOperator.GT, "5000");
        PatternFormatting patternFmt = rule.createPatternFormatting();
        patternFmt.setFillBackgroundColor(new XSSFColor(new byte[]{(byte) 255, (byte) 192, (byte) 0}, null));
        CellRangeAddress[] regions = {new CellRangeAddress(2, SUMMARY_ROWS - 1, 1, 1)};
        cf.addConditionalFormatting(regions, rule);

        // Column widths
        sheet.setColumnWidth(0, 20 * 256);
        sheet.setColumnWidth(1, 15 * 256);
        sheet.setColumnWidth(2, 15 * 256);
        sheet.setColumnWidth(3, 15 * 256);
        sheet.setColumnWidth(4, 10 * 256);
    }

    private void generateSalesPoi(XSSFWorkbook wb, XSSFCellStyle headerStyle,
                                   XSSFCellStyle currencyStyle, XSSFCellStyle percentStyle,
                                   XSSFCellStyle dateStyle) {
        XSSFSheet sheet = wb.createSheet("Sales");

        // Headers
        String[] headers = {"OrderID", "Date", "Customer", "Product", "Quantity", "UnitPrice",
            "Discount", "Total", "Tax", "NetAmount", "Region", "Status", "ShipDate", "Payment", "Notes"};
        XSSFRow headerRow = sheet.createRow(0);
        for (int c = 0; c < headers.length; c++) {
            XSSFCell cell = headerRow.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(headerStyle);
        }

        // Data rows
        LocalDate baseDate = LocalDate.of(2024, 1, 1);
        for (int r = 1; r <= SALES_ROWS; r++) {
            XSSFRow row = sheet.createRow(r);

            row.createCell(0).setCellValue("ORD-" + String.format("%06d", r));

            XSSFCell dateCell = row.createCell(1);
            dateCell.setCellValue(java.sql.Date.valueOf(baseDate.plusDays(random.nextInt(365))));
            dateCell.setCellStyle(dateStyle);

            row.createCell(2).setCellValue("Customer " + (random.nextInt(1000) + 1));
            row.createCell(3).setCellValue("Product " + (random.nextInt(500) + 1));
            row.createCell(4).setCellValue(random.nextInt(100) + 1);

            XSSFCell priceCell = row.createCell(5);
            priceCell.setCellValue(Math.round(random.nextDouble() * 500 * 100) / 100.0);
            priceCell.setCellStyle(currencyStyle);

            XSSFCell discountCell = row.createCell(6);
            discountCell.setCellValue(Math.round(random.nextDouble() * 0.3 * 100) / 100.0);
            discountCell.setCellStyle(percentStyle);

            // Formula columns
            int excelRow = r + 1;
            XSSFCell totalCell = row.createCell(7);
            totalCell.setCellFormula("E" + excelRow + "*F" + excelRow + "*(1-G" + excelRow + ")");
            totalCell.setCellStyle(currencyStyle);

            XSSFCell taxCell = row.createCell(8);
            taxCell.setCellFormula("H" + excelRow + "*0.08");
            taxCell.setCellStyle(currencyStyle);

            XSSFCell netCell = row.createCell(9);
            netCell.setCellFormula("H" + excelRow + "+I" + excelRow);
            netCell.setCellStyle(currencyStyle);

            row.createCell(10).setCellValue(REGIONS[random.nextInt(REGIONS.length)]);
            row.createCell(11).setCellValue(STATUSES[random.nextInt(STATUSES.length)]);

            XSSFCell shipCell = row.createCell(12);
            shipCell.setCellValue(java.sql.Date.valueOf(baseDate.plusDays(random.nextInt(365) + 1)));
            shipCell.setCellStyle(dateStyle);

            row.createCell(13).setCellValue(PAYMENTS[random.nextInt(PAYMENTS.length)]);
            row.createCell(14).setCellValue("Note for order " + r);
        }

        // AutoFilter
        sheet.setAutoFilter(new CellRangeAddress(0, SALES_ROWS, 0, SALES_COLS - 1));

        // Column widths
        double[] widths = {12, 12, 15, 15, 10, 12, 10, 12, 10, 12, 10, 10, 12, 15, 20};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, (int) (widths[c] * 256));
        }
    }

    private void generateCustomersPoi(XSSFWorkbook wb, XSSFCellStyle headerStyle,
                                       XSSFCellStyle dateStyle, XSSFCellStyle borderStyle) {
        XSSFSheet sheet = wb.createSheet("Customers");

        // Headers
        String[] headers = {"ID", "Name", "Email", "Phone", "Address", "City",
            "Country", "JoinDate", "Grade", "CreditLimit", "IsActive", "Memo"};
        XSSFRow headerRow = sheet.createRow(0);
        for (int c = 0; c < headers.length; c++) {
            XSSFCell cell = headerRow.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(headerStyle);
        }

        // Data rows
        LocalDate baseDate = LocalDate.of(2020, 1, 1);
        for (int r = 1; r <= CUSTOMERS_ROWS; r++) {
            XSSFRow row = sheet.createRow(r);

            XSSFCell idCell = row.createCell(0);
            idCell.setCellValue("CUST-" + String.format("%05d", r));
            idCell.setCellStyle(borderStyle);

            XSSFCell nameCell = row.createCell(1);
            nameCell.setCellValue("Customer Name " + r);
            nameCell.setCellStyle(borderStyle);

            XSSFCell emailCell = row.createCell(2);
            emailCell.setCellValue("customer" + r + "@example.com");
            emailCell.setCellStyle(borderStyle);

            XSSFCell phoneCell = row.createCell(3);
            phoneCell.setCellValue("+1-555-" + String.format("%04d", random.nextInt(10000)));
            phoneCell.setCellStyle(borderStyle);

            XSSFCell addrCell = row.createCell(4);
            addrCell.setCellValue(random.nextInt(9999) + " Main Street");
            addrCell.setCellStyle(borderStyle);

            XSSFCell cityCell = row.createCell(5);
            cityCell.setCellValue("City " + (random.nextInt(100) + 1));
            cityCell.setCellStyle(borderStyle);

            XSSFCell countryCell = row.createCell(6);
            countryCell.setCellValue(COUNTRIES[random.nextInt(COUNTRIES.length)]);
            countryCell.setCellStyle(borderStyle);

            XSSFCell joinCell = row.createCell(7);
            joinCell.setCellValue(java.sql.Date.valueOf(baseDate.plusDays(random.nextInt(1500))));
            joinCell.setCellStyle(dateStyle);

            XSSFCell gradeCell = row.createCell(8);
            gradeCell.setCellValue(GRADES[random.nextInt(GRADES.length)]);
            gradeCell.setCellStyle(borderStyle);

            XSSFCell creditCell = row.createCell(9);
            creditCell.setCellValue(Math.round(random.nextDouble() * 50000 * 100) / 100.0);
            creditCell.setCellStyle(borderStyle);

            XSSFCell activeCell = row.createCell(10);
            activeCell.setCellValue(random.nextBoolean());
            activeCell.setCellStyle(borderStyle);

            XSSFCell memoCell = row.createCell(11);
            memoCell.setCellValue("Memo for customer " + r);
            memoCell.setCellStyle(borderStyle);
        }

        // Data validation for Grade column
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
            dvHelper.createExplicitListConstraint(GRADES);
        CellRangeAddressList addressList = new CellRangeAddressList(1, CUSTOMERS_ROWS, 8, 8);
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
            dvConstraint, addressList);
        validation.setSuppressDropDownArrow(false);  // false = show dropdown
        sheet.addValidationData(validation);

        // Column widths
        double[] widths = {12, 20, 25, 15, 20, 12, 12, 12, 8, 12, 10, 25};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, (int) (widths[c] * 256));
        }
    }

    private void generateProductsPoi(XSSFWorkbook wb, XSSFCellStyle headerStyle,
                                      XSSFCellStyle currencyStyle, XSSFCellStyle percentStyle,
                                      XSSFCellStyle[] categoryStyles) {
        XSSFSheet sheet = wb.createSheet("Products");

        // Headers
        String[] headers = {"SKU", "Name", "Category", "Price", "Cost",
            "Margin", "Stock", "MinStock", "Reorder", "LastUpdate"};
        XSSFRow headerRow = sheet.createRow(0);
        for (int c = 0; c < headers.length; c++) {
            XSSFCell cell = headerRow.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(headerStyle);
        }

        // Data rows
        LocalDate baseDate = LocalDate.of(2024, 1, 1);
        for (int r = 1; r <= PRODUCTS_ROWS; r++) {
            XSSFRow row = sheet.createRow(r);
            int categoryIdx = random.nextInt(CATEGORIES.length);
            XSSFCellStyle categoryStyle = categoryStyles[categoryIdx];

            XSSFCell skuCell = row.createCell(0);
            skuCell.setCellValue("SKU-" + String.format("%06d", r));
            skuCell.setCellStyle(categoryStyle);

            XSSFCell nameCell = row.createCell(1);
            nameCell.setCellValue("Product Name " + r);
            nameCell.setCellStyle(categoryStyle);

            XSSFCell catCell = row.createCell(2);
            catCell.setCellValue(CATEGORIES[categoryIdx]);
            catCell.setCellStyle(categoryStyle);

            double price = Math.round(random.nextDouble() * 1000 * 100) / 100.0;
            double cost = Math.round(price * (0.4 + random.nextDouble() * 0.3) * 100) / 100.0;

            XSSFCell priceCell = row.createCell(3);
            priceCell.setCellValue(price);
            priceCell.setCellStyle(currencyStyle);

            XSSFCell costCell = row.createCell(4);
            costCell.setCellValue(cost);
            costCell.setCellStyle(currencyStyle);

            // Margin formula
            int excelRow = r + 1;
            XSSFCell marginCell = row.createCell(5);
            marginCell.setCellFormula("(D" + excelRow + "-E" + excelRow + ")/D" + excelRow);
            marginCell.setCellStyle(percentStyle);

            int stock = random.nextInt(1000);
            int minStock = random.nextInt(50) + 10;

            XSSFCell stockCell = row.createCell(6);
            stockCell.setCellValue(stock);
            stockCell.setCellStyle(categoryStyle);

            XSSFCell minStockCell = row.createCell(7);
            minStockCell.setCellValue(minStock);
            minStockCell.setCellStyle(categoryStyle);

            // Reorder formula (IF formula)
            XSSFCell reorderCell = row.createCell(8);
            reorderCell.setCellFormula("IF(G" + excelRow + "<H" + excelRow + ",\"Yes\",\"No\")");
            reorderCell.setCellStyle(categoryStyle);

            XSSFCell updateCell = row.createCell(9);
            updateCell.setCellValue(java.sql.Date.valueOf(baseDate.plusDays(random.nextInt(365))));
            updateCell.setCellStyle(categoryStyle);
        }

        // Column widths
        double[] widths = {15, 25, 12, 12, 12, 10, 10, 10, 10, 15};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, (int) (widths[c] * 256));
        }
    }

    private void generateConfigPoi(XSSFWorkbook wb, XSSFCellStyle headerStyle, XSSFCellStyle borderStyle) {
        XSSFSheet sheet = wb.createSheet("Config");

        // Headers
        String[] headers = {"Key", "Value", "Description"};
        XSSFRow headerRow = sheet.createRow(0);
        for (int c = 0; c < headers.length; c++) {
            XSSFCell cell = headerRow.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(headerStyle);
        }

        // Configuration data with various types
        int rowNum = 1;

        // Boolean values
        XSSFRow row1 = sheet.createRow(rowNum);
        XSSFCell key1 = row1.createCell(0);
        key1.setCellValue("EnableFeatureA");
        key1.setCellStyle(borderStyle);
        XSSFCell val1 = row1.createCell(1);
        val1.setCellValue(true);
        val1.setCellStyle(borderStyle);
        XSSFCell desc1 = row1.createCell(2);
        desc1.setCellValue("Enable feature A");
        desc1.setCellStyle(borderStyle);
        rowNum++;

        XSSFRow row2 = sheet.createRow(rowNum);
        XSSFCell key2 = row2.createCell(0);
        key2.setCellValue("EnableFeatureB");
        key2.setCellStyle(borderStyle);
        XSSFCell val2 = row2.createCell(1);
        val2.setCellValue(false);
        val2.setCellStyle(borderStyle);
        XSSFCell desc2 = row2.createCell(2);
        desc2.setCellValue("Enable feature B");
        desc2.setCellStyle(borderStyle);
        rowNum++;

        // Error values
        XSSFRow row3 = sheet.createRow(rowNum);
        XSSFCell key3 = row3.createCell(0);
        key3.setCellValue("ErrorNA");
        key3.setCellStyle(borderStyle);
        XSSFCell val3 = row3.createCell(1);
        val3.setCellErrorValue(FormulaError.NA);
        val3.setCellStyle(borderStyle);
        XSSFCell desc3 = row3.createCell(2);
        desc3.setCellValue("Not available error");
        desc3.setCellStyle(borderStyle);
        rowNum++;

        XSSFRow row4 = sheet.createRow(rowNum);
        XSSFCell key4 = row4.createCell(0);
        key4.setCellValue("ErrorREF");
        key4.setCellStyle(borderStyle);
        XSSFCell val4 = row4.createCell(1);
        val4.setCellErrorValue(FormulaError.REF);
        val4.setCellStyle(borderStyle);
        XSSFCell desc4 = row4.createCell(2);
        desc4.setCellValue("Reference error");
        desc4.setCellStyle(borderStyle);
        rowNum++;

        XSSFRow row5 = sheet.createRow(rowNum);
        XSSFCell key5 = row5.createCell(0);
        key5.setCellValue("ErrorVALUE");
        key5.setCellStyle(borderStyle);
        XSSFCell val5 = row5.createCell(1);
        val5.setCellErrorValue(FormulaError.VALUE);
        val5.setCellStyle(borderStyle);
        XSSFCell desc5 = row5.createCell(2);
        desc5.setCellValue("Value error");
        desc5.setCellStyle(borderStyle);
        rowNum++;

        // Fill remaining rows
        for (; rowNum < CONFIG_ROWS; rowNum++) {
            String key = "Config_" + (rowNum - 5);
            XSSFRow row = sheet.createRow(rowNum);

            XSSFCell keyCell = row.createCell(0);
            keyCell.setCellValue(key);
            keyCell.setCellStyle(borderStyle);

            XSSFCell valCell = row.createCell(1);
            // Alternate between different value types
            switch (rowNum % 4) {
                case 0 -> valCell.setCellValue("StringValue" + rowNum);
                case 1 -> valCell.setCellValue(random.nextDouble() * 1000);
                case 2 -> valCell.setCellValue(random.nextBoolean());
                case 3 -> valCell.setCellValue(random.nextInt(1000));
            }
            valCell.setCellStyle(borderStyle);

            XSSFCell descCell = row.createCell(2);
            descCell.setCellValue("Description for " + key);
            descCell.setCellStyle(borderStyle);
        }

        // Sheet protection
        sheet.protectSheet("");

        // Column widths
        sheet.setColumnWidth(0, 20 * 256);
        sheet.setColumnWidth(1, 20 * 256);
        sheet.setColumnWidth(2, 40 * 256);
    }

    /**
     * Calculates the total number of cells generated.
     */
    public long calculateTotalCells() {
        return (long) SUMMARY_ROWS * SUMMARY_COLS
            + (long) (SALES_ROWS + 1) * SALES_COLS        // +1 for header
            + (long) (CUSTOMERS_ROWS + 1) * CUSTOMERS_COLS
            + (long) (PRODUCTS_ROWS + 1) * PRODUCTS_COLS
            + (long) CONFIG_ROWS * CONFIG_COLS;
    }

    /**
     * Resets the random number generator with the fixed seed.
     */
    public void resetRandom() {
        random = new Random(42);
    }
}
