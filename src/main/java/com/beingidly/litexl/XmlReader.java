package com.beingidly.litexl;

import org.jspecify.annotations.Nullable;

import javax.xml.stream.*;
import java.io.InputStream;

/**
 * SAX-style XML reader wrapper using StAX.
 */
public final class XmlReader implements AutoCloseable {

    private final XMLStreamReader reader;

    /** XML event types. */
    public enum Event {
        /** Start of an element. */
        START_ELEMENT,
        /** End of an element. */
        END_ELEMENT,
        /** Character data. */
        CHARACTERS,
        /** End of document. */
        END_DOCUMENT
    }

    /**
     * Creates a new XML reader from the given input stream.
     *
     * @param input the input stream
     */
    public XmlReader(InputStream input) {
        try {
            XMLInputFactory factory = XMLInputFactory.newInstance();
            // Security: disable external entities
            factory.setProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES, false);
            factory.setProperty(XMLInputFactory.SUPPORT_DTD, false);
            this.reader = factory.createXMLStreamReader(input);
        } catch (XMLStreamException e) {
            throw new RuntimeException("Failed to create XML reader", e);
        }
    }

    /**
     * Returns true if there are more XML events to read.
     *
     * @return true if more events available
     */
    public boolean hasNext() {
        try {
            return reader.hasNext();
        } catch (XMLStreamException e) {
            throw new RuntimeException("XML read error", e);
        }
    }

    /**
     * Returns the next XML event.
     *
     * @return the next event
     */
    public Event next() {
        try {
            while (reader.hasNext()) {
                int event = reader.next();
                switch (event) {
                    case XMLStreamConstants.START_ELEMENT:
                        return Event.START_ELEMENT;
                    case XMLStreamConstants.END_ELEMENT:
                        return Event.END_ELEMENT;
                    case XMLStreamConstants.CHARACTERS:
                        if (!reader.isWhiteSpace()) {
                            return Event.CHARACTERS;
                        }
                        break;
                    case XMLStreamConstants.END_DOCUMENT:
                        return Event.END_DOCUMENT;
                }
            }
            return Event.END_DOCUMENT;
        } catch (XMLStreamException e) {
            throw new RuntimeException("XML read error", e);
        }
    }

    /**
     * Returns the local name of the current element.
     *
     * @return the local name
     */
    public String getLocalName() {
        return reader.getLocalName();
    }

    /**
     * Returns the value of the given attribute, or null if not found.
     *
     * @param name the attribute name
     * @return the attribute value, or null
     */
    public @Nullable String getAttributeValue(String name) {
        // Try without namespace first
        String value = reader.getAttributeValue(null, name);
        if (value != null) {
            return value;
        }

        // Try with common namespaces
        value = reader.getAttributeValue("http://schemas.openxmlformats.org/officeDocument/2006/relationships", name.replace("r:", ""));
        return value;
    }

    /**
     * Returns the text content of the current element.
     *
     * @return the element text
     */
    public String getElementText() {
        try {
            return reader.getElementText();
        } catch (XMLStreamException e) {
            throw new RuntimeException("XML read error", e);
        }
    }

    /**
     * Returns the text of the current CHARACTERS event.
     *
     * @return the text content
     */
    public String getText() {
        return reader.getText();
    }

    @Override
    public void close() {
        try {
            reader.close();
        } catch (XMLStreamException e) {
            // Ignore close errors
        }
    }
}
