package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import static org.junit.jupiter.api.Assertions.*;

class XmlReaderTest {

    @Test
    void readsSimpleElement() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root><child>text</child></root>";
        try (var reader = new XmlReader(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)))) {
            assertTrue(reader.hasNext());
            while (reader.hasNext()) {
                var event = reader.next();
                if (event == XmlReader.Event.START_ELEMENT && "child".equals(reader.getLocalName())) {
                    assertEquals("text", reader.getElementText());
                    break;
                }
            }
        }
    }

    @Test
    void readsAttributes() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root attr=\"value\"/>";
        try (var reader = new XmlReader(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)))) {
            while (reader.hasNext()) {
                var event = reader.next();
                if (event == XmlReader.Event.START_ELEMENT && "root".equals(reader.getLocalName())) {
                    assertEquals("value", reader.getAttributeValue("attr"));
                    break;
                }
            }
        }
    }

    @Test
    void returnsNullForMissingAttribute() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root/>";
        try (var reader = new XmlReader(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)))) {
            while (reader.hasNext()) {
                var event = reader.next();
                if (event == XmlReader.Event.START_ELEMENT) {
                    assertNull(reader.getAttributeValue("nonexistent"));
                    break;
                }
            }
        }
    }

    @Test
    void skipsWhitespaceOnlyText() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root>  \n  <child/></root>";
        try (var reader = new XmlReader(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)))) {
            int charEvents = 0;
            while (reader.hasNext()) {
                var event = reader.next();
                if (event == XmlReader.Event.CHARACTERS) {
                    charEvents++;
                }
            }
            assertEquals(0, charEvents);
        }
    }

    @Test
    void handlesEmptyDocument() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root/>";
        try (var reader = new XmlReader(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)))) {
            assertTrue(reader.hasNext());
        }
    }

    @Test
    void handlesNamespacedAttributes() throws Exception {
        String xml = "<?xml version=\"1.0\"?>" +
            "<root xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
            "r:id=\"rId1\"/>";
        try (var reader = new XmlReader(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)))) {
            while (reader.hasNext()) {
                var event = reader.next();
                if (event == XmlReader.Event.START_ELEMENT) {
                    String value = reader.getAttributeValue("id");
                    assertNotNull(value);
                    assertEquals("rId1", value);
                    break;
                }
            }
        }
    }

    @Test
    void closesCleanly() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root/>";
        var reader = new XmlReader(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)));
        reader.close();
    }

    @Test
    void parseEmptyElement() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root><empty/></root>";
        try (var is = new ByteArrayInputStream(xml.getBytes());
             var reader = new XmlReader(is)) {
            assertEquals(XmlReader.Event.START_ELEMENT, reader.next());
            assertEquals("root", reader.getLocalName());
            assertEquals(XmlReader.Event.START_ELEMENT, reader.next());
            assertEquals("empty", reader.getLocalName());
            assertEquals(XmlReader.Event.END_ELEMENT, reader.next());
            assertEquals(XmlReader.Event.END_ELEMENT, reader.next());
        }
    }

    @Test
    void parseWithAttributes() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root attr1=\"value1\" attr2=\"value2\"/>";
        try (var is = new ByteArrayInputStream(xml.getBytes());
             var reader = new XmlReader(is)) {
            assertEquals(XmlReader.Event.START_ELEMENT, reader.next());
            assertEquals("value1", reader.getAttributeValue("attr1"));
            assertEquals("value2", reader.getAttributeValue("attr2"));
            assertNull(reader.getAttributeValue("nonexistent"));
        }
    }

    @Test
    void parseText() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root>Hello World</root>";
        try (var is = new ByteArrayInputStream(xml.getBytes());
             var reader = new XmlReader(is)) {
            assertEquals(XmlReader.Event.START_ELEMENT, reader.next());
            assertEquals(XmlReader.Event.CHARACTERS, reader.next());
            assertEquals("Hello World", reader.getText());
            assertEquals(XmlReader.Event.END_ELEMENT, reader.next());
        }
    }

    @Test
    void parseNestedElements() throws Exception {
        String xml = "<?xml version=\"1.0\"?><root><child><grandchild/></child></root>";
        try (var is = new ByteArrayInputStream(xml.getBytes());
             var reader = new XmlReader(is)) {
            assertEquals(XmlReader.Event.START_ELEMENT, reader.next());
            assertEquals("root", reader.getLocalName());
            assertEquals(XmlReader.Event.START_ELEMENT, reader.next());
            assertEquals("child", reader.getLocalName());
            assertEquals(XmlReader.Event.START_ELEMENT, reader.next());
            assertEquals("grandchild", reader.getLocalName());
        }
    }

    @Test
    void eventValues() {
        assertEquals(4, XmlReader.Event.values().length);
        assertNotNull(XmlReader.Event.valueOf("START_ELEMENT"));
        assertNotNull(XmlReader.Event.valueOf("END_ELEMENT"));
        assertNotNull(XmlReader.Event.valueOf("CHARACTERS"));
        assertNotNull(XmlReader.Event.valueOf("END_DOCUMENT"));
    }
}
