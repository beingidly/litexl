package com.beingidly.litexl.spi;

import org.jspecify.annotations.Nullable;

import java.io.IOException;
import java.io.InputStream;
import java.util.Set;

/**
 * Context for read extensions to access ZIP entries.
 */
public interface ReadContext {

    /**
     * Opens a ZIP entry for reading.
     *
     * @param path the entry path within the ZIP
     * @return an input stream, or null if the entry does not exist
     */
    @Nullable InputStream openEntry(String path) throws IOException;

    /**
     * Checks if a ZIP entry exists.
     *
     * @param path the entry path within the ZIP
     * @return true if the entry exists
     */
    boolean hasEntry(String path);

    /**
     * Returns all entry names in the ZIP.
     */
    Set<String> entryNames();
}
