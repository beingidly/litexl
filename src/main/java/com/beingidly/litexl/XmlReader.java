package com.beingidly.litexl;

import org.jspecify.annotations.Nullable;

import javax.xml.stream.*;
import java.io.InputStream;

/**
 * SAX-style XML reader wrapper using StAX.
 */
final class XmlReader implements AutoCloseable {

    private final XMLStreamReader reader;

    public enum Event {
        START_ELEMENT,
        END_ELEMENT,
        CHARACTERS,
        END_DOCUMENT
    }

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

    public boolean hasNext() {
        try {
            return reader.hasNext();
        } catch (XMLStreamException e) {
            throw new RuntimeException("XML read error", e);
        }
    }

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

    public String getLocalName() {
        return reader.getLocalName();
    }

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

    public String getElementText() {
        try {
            return reader.getElementText();
        } catch (XMLStreamException e) {
            throw new RuntimeException("XML read error", e);
        }
    }

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
