package com.beingidly.litexl;

import javax.xml.stream.*;
import java.io.OutputStream;
import java.io.IOException;

/**
 * Streaming XML writer wrapper using StAX.
 */
public final class XmlWriter implements AutoCloseable {

    private final XMLStreamWriter writer;

    /**
     * Creates a new XML writer for the given output stream.
     *
     * @param output the output stream
     * @throws IOException if the writer cannot be created
     */
    public XmlWriter(OutputStream output) throws IOException {
        try {
            XMLOutputFactory factory = XMLOutputFactory.newInstance();
            this.writer = factory.createXMLStreamWriter(output, "UTF-8");
        } catch (XMLStreamException e) {
            throw new IOException("Failed to create XML writer", e);
        }
    }

    /**
     * Writes the XML declaration.
     *
     * @throws IOException if an I/O error occurs
     */
    public void startDocument() throws IOException {
        try {
            writer.writeStartDocument("UTF-8", "1.0");
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Ends the document and flushes.
     *
     * @throws IOException if an I/O error occurs
     */
    public void endDocument() throws IOException {
        try {
            writer.writeEndDocument();
            writer.flush();
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Starts an element with the given name.
     *
     * @param name the element name
     * @throws IOException if an I/O error occurs
     */
    public void startElement(String name) throws IOException {
        try {
            writer.writeStartElement(name);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Starts a namespace-qualified element with proper namespace binding.
     *
     * @param prefix       the namespace prefix (e.g. "c")
     * @param localName    the local name (e.g. "chartSpace")
     * @param namespaceUri the namespace URI
     * @throws IOException if an I/O error occurs
     */
    public void startElement(String prefix, String localName, String namespaceUri) throws IOException {
        try {
            writer.writeStartElement(prefix, localName, namespaceUri);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Ends the current element.
     *
     * @throws IOException if an I/O error occurs
     */
    public void endElement() throws IOException {
        try {
            writer.writeEndElement();
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Writes an empty element.
     *
     * @param name the element name
     * @throws IOException if an I/O error occurs
     */
    public void emptyElement(String name) throws IOException {
        try {
            writer.writeEmptyElement(name);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Writes an empty namespace-qualified element with proper namespace binding.
     *
     * @param prefix       the namespace prefix (e.g. "c")
     * @param localName    the local name (e.g. "layout")
     * @param namespaceUri the namespace URI
     * @throws IOException if an I/O error occurs
     */
    public void emptyElement(String prefix, String localName, String namespaceUri) throws IOException {
        try {
            writer.writeEmptyElement(prefix, localName, namespaceUri);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Writes an attribute on the current element.
     *
     * @param name the attribute name
     * @param value the attribute value
     * @throws IOException if an I/O error occurs
     */
    public void attribute(String name, String value) throws IOException {
        try {
            writer.writeAttribute(name, value);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Writes a namespace declaration (xmlns:prefix="uri").
     *
     * @param prefix       the namespace prefix
     * @param namespaceUri the namespace URI
     * @throws IOException if an I/O error occurs
     */
    public void namespace(String prefix, String namespaceUri) throws IOException {
        try {
            writer.writeNamespace(prefix, namespaceUri);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Sets a namespace prefix binding for subsequent elements.
     *
     * @param prefix       the namespace prefix
     * @param namespaceUri the namespace URI
     * @throws IOException if an I/O error occurs
     */
    public void setPrefix(String prefix, String namespaceUri) throws IOException {
        try {
            writer.setPrefix(prefix, namespaceUri);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    /**
     * Writes text content.
     *
     * @param text the text to write
     * @throws IOException if an I/O error occurs
     */
    public void text(String text) throws IOException {
        try {
            writer.writeCharacters(text);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    @Override
    public void close() throws IOException {
        try {
            writer.close();
        } catch (XMLStreamException e) {
            throw new IOException("Failed to close XML writer", e);
        }
    }
}
