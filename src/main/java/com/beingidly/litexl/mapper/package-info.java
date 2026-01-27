/**
 * Annotation-based mapping between Java classes and Excel workbooks.
 *
 * <h2>Quick Start</h2>
 * <pre>{@code
 * @LitexlRow
 * record Person(
 *     @LitexlColumn(index = 0, header = "Name") String name,
 *     @LitexlColumn(index = 1, header = "Age") int age
 * ) {}
 *
 * @LitexlWorkbook
 * record Report(
 *     @LitexlSheet(name = "People") List<Person> people
 * ) {}
 *
 * // Write
 * LitexlMapper.write(report, Path.of("output.xlsx"));
 *
 * // Read
 * Report result = LitexlMapper.read(Path.of("input.xlsx"), Report.class);
 * }</pre>
 *
 * @see com.beingidly.litexl.mapper.LitexlMapper
 * @see com.beingidly.litexl.mapper.LitexlWorkbook
 * @see com.beingidly.litexl.mapper.LitexlSheet
 * @see com.beingidly.litexl.mapper.LitexlRow
 * @see com.beingidly.litexl.mapper.LitexlColumn
 */
@NullMarked
package com.beingidly.litexl.mapper;

import org.jspecify.annotations.NullMarked;
