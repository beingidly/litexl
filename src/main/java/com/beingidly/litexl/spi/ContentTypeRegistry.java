package com.beingidly.litexl.spi;

import java.io.IOException;

/**
 * Registry for contributing additional content types to [Content_Types].xml.
 */
public interface ContentTypeRegistry {

    /**
     * Registers a Default content type by file extension.
     *
     * @param extension   the file extension (e.g. "png")
     * @param contentType the MIME content type
     */
    void addDefault(String extension, String contentType) throws IOException;

    /**
     * Registers an Override content type by part name.
     *
     * @param partName    the part name (e.g. "/xl/charts/chart1.xml")
     * @param contentType the MIME content type
     */
    void addOverride(String partName, String contentType) throws IOException;
}
