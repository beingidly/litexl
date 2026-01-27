@NullMarked
/**
 * Lightweight Excel library for reading and writing XLSX files.
 *
 * <p>The main entry points are:</p>
 * <ul>
 *   <li>{@link com.beingidly.litexl.Workbook} - The root container for Excel data</li>
 *   <li>{@link com.beingidly.litexl.Sheet} - A worksheet within a workbook</li>
 *   <li>{@link com.beingidly.litexl.Row} - A row within a worksheet</li>
 *   <li>{@link com.beingidly.litexl.Cell} - A cell within a row</li>
 * </ul>
 *
 * <h2>Thread Safety</h2>
 * <p>Classes in this package are <b>not thread-safe</b> unless explicitly
 * documented otherwise. Use external synchronization for concurrent access
 * to shared instances.</p>
 *
 * <h2>Example Usage</h2>
 * <pre>{@code
 * // Create a new workbook
 * try (Workbook wb = Workbook.create()) {
 *     Sheet sheet = wb.addSheet("Data");
 *     sheet.cell(0, 0).set("Hello, World!");
 *     wb.save(Path.of("output.xlsx"));
 * }
 *
 * // Read an existing workbook
 * try (Workbook wb = Workbook.open(Path.of("input.xlsx"))) {
 *     Sheet sheet = wb.getSheet(0);
 *     String value = sheet.cell(0, 0).string();
 * }
 * }</pre>
 *
 * @see com.beingidly.litexl.Workbook
 */
package com.beingidly.litexl;

import org.jspecify.annotations.NullMarked;
