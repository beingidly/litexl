package com.beingidly.litexl.chart.internal;

import com.beingidly.litexl.XmlReader;
import com.beingidly.litexl.chart.ChartPosition;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Reads drawing XML (xl/drawings/drawingN.xml) to find chart references and positions.
 */
final class DrawingXmlReader {

    record ChartRef(String rId, ChartPosition position) {}

    static List<ChartRef> read(InputStream is) throws IOException {
        List<ChartRef> refs = new ArrayList<>();
        try (XmlReader xml = new XmlReader(is)) {
            int fromCol = 0, fromRow = 0, fromColOff = 0, fromRowOff = 0;
            int toCol = 0, toRow = 0, toColOff = 0, toRowOff = 0;
            boolean inFrom = false, inTo = false;
            String currentRId = null;
            String lastLocalName = null;

            while (xml.hasNext()) {
                XmlReader.Event event = xml.next();
                if (event == XmlReader.Event.START_ELEMENT) {
                    String name = xml.getLocalName();
                    lastLocalName = name;

                    switch (name) {
                        case "from" -> inFrom = true;
                        case "to" -> inTo = true;
                        case "chart" -> {
                            String rId = xml.getAttributeValue("r:id");
                            if (rId == null) {
                                rId = xml.getAttributeValue("id");
                            }
                            currentRId = rId;
                        }
                    }
                } else if (event == XmlReader.Event.END_ELEMENT) {
                    String name = xml.getLocalName();
                    switch (name) {
                        case "from" -> inFrom = false;
                        case "to" -> inTo = false;
                        case "twoCellAnchor" -> {
                            if (currentRId != null) {
                                refs.add(new ChartRef(currentRId,
                                    new ChartPosition(fromCol, fromRow, fromColOff, fromRowOff,
                                        toCol, toRow, toColOff, toRowOff)));
                            }
                            currentRId = null;
                            fromCol = fromRow = fromColOff = fromRowOff = 0;
                            toCol = toRow = toColOff = toRowOff = 0;
                        }
                    }
                } else if (event == XmlReader.Event.CHARACTERS && lastLocalName != null) {
                    String text = xml.getText().trim();
                    if (text.isEmpty()) continue;

                    try {
                        int val = Integer.parseInt(text);
                        if (inFrom) {
                            switch (lastLocalName) {
                                case "col" -> fromCol = val;
                                case "row" -> fromRow = val;
                                case "colOff" -> fromColOff = val;
                                case "rowOff" -> fromRowOff = val;
                            }
                        } else if (inTo) {
                            switch (lastLocalName) {
                                case "col" -> toCol = val;
                                case "row" -> toRow = val;
                                case "colOff" -> toColOff = val;
                                case "rowOff" -> toRowOff = val;
                            }
                        }
                    } catch (NumberFormatException _) {
                        // ignore
                    }
                    lastLocalName = null;
                }
            }
        }
        return refs;
    }

    private DrawingXmlReader() {}
}
