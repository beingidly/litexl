package com.beingidly.litexl.chart.internal;

import com.beingidly.litexl.XmlWriter;
import com.beingidly.litexl.chart.Chart;
import com.beingidly.litexl.chart.ChartPosition;

import java.io.IOException;
import java.util.List;

/**
 * Writes drawing XML (xl/drawings/drawingN.xml).
 *
 * <p>Uses proper StAX namespace bindings for Excel compatibility.
 */
final class DrawingXmlWriter {

    private static final String NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static final String NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static final int DEFAULT_CX = 4572000;
    private static final int DEFAULT_CY = 2743200;

    static void write(XmlWriter xml, List<Chart> charts, int firstChartNum) throws IOException {
        xml.startDocument();

        xml.setPrefix("xdr", NS_XDR);
        xml.setPrefix("a", NS_A);
        xml.setPrefix("c", NS_C);
        xml.setPrefix("r", NS_R);
        xml.startElement("xdr", "wsDr", NS_XDR);
        xml.namespace("xdr", NS_XDR);
        xml.namespace("a", NS_A);
        xml.namespace("c", NS_C);
        xml.namespace("r", NS_R);

        for (int i = 0; i < charts.size(); i++) {
            writeTwoCellAnchor(xml, charts.get(i).position(), "rId" + (i + 1), i);
        }

        xml.endElement(); // xdr:wsDr
        xml.endDocument();
    }

    private static void writeTwoCellAnchor(XmlWriter xml, ChartPosition pos,
                                            String rId, int chartIndex) throws IOException {
        xml.startElement("xdr", "twoCellAnchor", NS_XDR);
        xml.attribute("editAs", "twoCell");

        // From
        xml.startElement("xdr", "from", NS_XDR);
        xdrIntElement(xml, "col", pos.fromCol());
        xdrIntElement(xml, "colOff", pos.fromColOff());
        xdrIntElement(xml, "row", pos.fromRow());
        xdrIntElement(xml, "rowOff", pos.fromRowOff());
        xml.endElement();

        // To
        xml.startElement("xdr", "to", NS_XDR);
        xdrIntElement(xml, "col", pos.toCol());
        xdrIntElement(xml, "colOff", pos.toColOff());
        xdrIntElement(xml, "row", pos.toRow());
        xdrIntElement(xml, "rowOff", pos.toRowOff());
        xml.endElement();

        // Graphic frame
        xml.startElement("xdr", "graphicFrame", NS_XDR);

        xml.startElement("xdr", "nvGraphicFramePr", NS_XDR);
        xml.emptyElement("xdr", "cNvPr", NS_XDR);
        xml.attribute("id", String.valueOf(chartIndex));
        xml.attribute("name", "Chart " + chartIndex);
        xml.emptyElement("xdr", "cNvGraphicFramePr", NS_XDR);
        xml.endElement(); // nvGraphicFramePr

        xml.startElement("xdr", "xfrm", NS_XDR);
        xml.emptyElement("a", "off", NS_A);
        xml.attribute("x", "0");
        xml.attribute("y", "0");
        xml.emptyElement("a", "ext", NS_A);
        xml.attribute("cx", String.valueOf(DEFAULT_CX));
        xml.attribute("cy", String.valueOf(DEFAULT_CY));
        xml.endElement(); // xfrm

        xml.startElement("a", "graphic", NS_A);
        xml.startElement("a", "graphicData", NS_A);
        xml.attribute("uri", NS_C);
        xml.emptyElement("c", "chart", NS_C);
        xml.attribute("r:id", rId);
        xml.endElement(); // graphicData
        xml.endElement(); // graphic

        xml.endElement(); // graphicFrame

        xml.emptyElement("xdr", "clientData", NS_XDR);

        xml.endElement(); // twoCellAnchor
    }

    private static void xdrIntElement(XmlWriter xml, String localName, int value) throws IOException {
        xml.startElement("xdr", localName, NS_XDR);
        xml.text(String.valueOf(value));
        xml.endElement();
    }

    private DrawingXmlWriter() {}
}
