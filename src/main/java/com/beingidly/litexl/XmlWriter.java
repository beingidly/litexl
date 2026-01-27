package com.beingidly.litexl;

import javax.xml.stream.*;
import java.io.OutputStream;
import java.io.IOException;

/**
 * Streaming XML writer wrapper using StAX.
 */
final class XmlWriter implements AutoCloseable {

    private final XMLStreamWriter writer;

    public XmlWriter(OutputStream output) throws IOException {
        try {
            XMLOutputFactory factory = XMLOutputFactory.newInstance();
            this.writer = factory.createXMLStreamWriter(output, "UTF-8");
        } catch (XMLStreamException e) {
            throw new IOException("Failed to create XML writer", e);
        }
    }

    public void startDocument() throws IOException {
        try {
            writer.writeStartDocument("UTF-8", "1.0");
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    public void endDocument() throws IOException {
        try {
            writer.writeEndDocument();
            writer.flush();
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    public void startElement(String name) throws IOException {
        try {
            writer.writeStartElement(name);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    public void endElement() throws IOException {
        try {
            writer.writeEndElement();
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    public void emptyElement(String name) throws IOException {
        try {
            writer.writeEmptyElement(name);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

    public void attribute(String name, String value) throws IOException {
        try {
            writer.writeAttribute(name, value);
        } catch (XMLStreamException e) {
            throw new IOException("XML write error", e);
        }
    }

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
