package com.beingidly.litexl.spi;

import com.beingidly.litexl.Workbook;

import java.io.IOException;

/**
 * SPI for extending the XLSX read process.
 *
 * <p>Implementations are discovered via {@link java.util.ServiceLoader}.
 * When {@code litexl-chart} (or another extension) is on the classpath,
 * its {@code ReadExtension} is automatically invoked after standard parts are read.
 */
public interface ReadExtension {

    /**
     * Called after all standard XLSX parts (shared strings, styles, workbook, sheets)
     * have been read. The extension can use the {@link ReadContext} to access additional
     * ZIP entries (e.g. chart XML, drawing XML) and attach parsed data to sheets.
     *
     * @param ctx      the read context for accessing ZIP entries
     * @param workbook the workbook being read
     * @throws IOException if an I/O error occurs
     */
    void read(ReadContext ctx, Workbook workbook) throws IOException;
}
