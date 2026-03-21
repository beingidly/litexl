package com.beingidly.litexl.spi;

import java.io.IOException;
import java.io.OutputStream;

/**
 * Context for write extensions to create additional ZIP entries.
 */
public interface WriteContext {

    /**
     * Opens a new ZIP entry and returns the output stream for writing.
     *
     * <p>The caller should close the returned stream when done writing.
     *
     * @param path the entry path within the ZIP (e.g. "xl/charts/chart1.xml")
     * @return an output stream for writing entry content
     * @throws IOException if an I/O error occurs
     */
    OutputStream newEntry(String path) throws IOException;
}
