package com.beingidly.litexl.chart.internal;

import com.beingidly.litexl.XmlWriter;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.axis.*;
import com.beingidly.litexl.chart.style.*;

import java.io.IOException;
import java.util.List;

import static com.beingidly.litexl.chart.ChartInternalAccess.*;
import static com.beingidly.litexl.chart.axis.AxisInternalAccess.*;
import static com.beingidly.litexl.chart.style.StyleInternalAccess.*;

/**
 * Writes chart XML (xl/charts/chartN.xml).
 *
 * <p>Uses proper StAX namespace bindings for Excel compatibility.
 */
final class ChartXmlWriter {

    static final String NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    static final String NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    static final String NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    static void write(XmlWriter xml, Chart chart, String sheetName) throws IOException {
        xml.startDocument();

        // Root element with proper namespace bindings
        xml.setPrefix("c", NS_C);
        xml.setPrefix("a", NS_A);
        xml.startElement("c", "chartSpace", NS_C);
        xml.namespace("c", NS_C);
        xml.namespace("a", NS_A);

        cStart(xml, "chart");

        // Title
        if (chart.title() != null) {
            writeTitle(xml, chart.title());
        } else {
            cEmptyVal(xml, "autoTitleDeleted", "1");
        }

        // 3D view
        if (chart.plotConfig().view3D() != null) {
            writeView3D(xml, chart.plotConfig().view3D());
        }

        // Plot area
        cStart(xml, "plotArea");
        cEmpty(xml, "layout");
        writeChartTypeElement(xml, chart, sheetName);
        if (chart.type().hasAxes()) {
            writeAxes(xml, chart);
        }
        xml.endElement(); // c:plotArea

        // Legend (optional)
        if (chart.legend() != null && chart.legend().position() != LegendPosition.NONE) {
            writeLegend(xml, chart.legend());
        }

        cEmptyVal(xml, "plotVisOnly", "true");

        xml.endElement(); // c:chart

        writePrintSettings(xml);

        xml.endElement(); // c:chartSpace
        xml.endDocument();
    }

    // --- Title ---

    private static void writeTitle(XmlWriter xml, ChartTitle title) throws IOException {
        cStart(xml, "title");
        cStart(xml, "tx");
        cStart(xml, "rich");
        aEmpty(xml, "bodyPr");
        xml.attribute("anchor", "t");
        xml.attribute("rtlCol", "false");
        aEmpty(xml, "lstStyle");
        aStart(xml, "p");
        aStart(xml, "pPr");
        xml.attribute("algn", "l");
        aEmpty(xml, "defRPr");
        xml.endElement(); // a:pPr

        aStart(xml, "r");
        if (title.font() != null) {
            writeRunProperties(xml, title.font());
        } else {
            aEmpty(xml, "rPr");
            xml.attribute("lang", "en-US");
        }
        aStart(xml, "t");
        xml.text(title.text());
        xml.endElement(); // a:t
        xml.endElement(); // a:r

        aEmpty(xml, "endParaRPr");
        xml.attribute("lang", "en-US");
        xml.attribute("sz", "1100");

        xml.endElement(); // a:p
        xml.endElement(); // c:rich
        xml.endElement(); // c:tx
        cEmpty(xml, "layout");
        xml.endElement(); // c:title
    }

    // --- 3D view ---

    private static void writeView3D(XmlWriter xml, ChartView3D view) throws IOException {
        cStart(xml, "view3D");
        cEmptyVal(xml, "rotX", String.valueOf(view.xRotation()));
        cEmptyVal(xml, "rotY", String.valueOf(view.yRotation()));
        cEmptyVal(xml, "perspective", String.valueOf(view.perspective()));
        xml.endElement();
    }

    // --- Chart type element ---

    private static void writeChartTypeElement(XmlWriter xml, Chart chart, String sheetName) throws IOException {
        // xmlTag returns "c:barChart" etc. - extract local name
        String tag = xmlTag(chart.type());
        String localName = tag.substring(tag.indexOf(':') + 1);
        xml.startElement("c", localName, NS_C);

        ChartPlotConfig config = chart.plotConfig();

        if (config.barDirection() != null) {
            cEmptyVal(xml, "barDir", xmlValue(config.barDirection()));
        }

        if (config.grouping() != null && config.grouping() != Grouping.CLUSTERED) {
            cEmptyVal(xml, "grouping", xmlValue(config.grouping()));
        }

        if (config.scatterStyle() != null) {
            cEmptyVal(xml, "scatterStyle", xmlValue(config.scatterStyle()));
        }

        if (config.radarStyle() != null) {
            cEmptyVal(xml, "radarStyle", xmlValue(config.radarStyle()));
        }

        // Series
        for (int i = 0; i < chart.series().size(); i++) {
            writeSeries(xml, chart.series().get(i), i, sheetName, chart.type());
        }

        // Axis IDs
        if (chart.type().hasAxes()) {
            cEmptyVal(xml, "axId", "0");
            cEmptyVal(xml, "axId", "1");
        }

        xml.endElement();
    }

    // --- Series ---

    private static void writeSeries(XmlWriter xml, ChartSeries series, int index,
                                     String sheetName, ChartType chartType) throws IOException {
        cStart(xml, "ser");

        cEmptyVal(xml, "idx", String.valueOf(index));
        cEmptyVal(xml, "order", String.valueOf(index));

        if (series.name() != null) {
            cStart(xml, "tx");
            cStart(xml, "v");
            xml.text(series.name());
            xml.endElement(); // c:v
            xml.endElement(); // c:tx
        }

        if (series.fill() != null || series.line() != null) {
            cStart(xml, "spPr");
            if (series.fill() != null) writeFill(xml, series.fill());
            if (series.line() != null) writeLine(xml, series.line());
            xml.endElement();
        }

        if (series.marker() != null) writeMarker(xml, series.marker());
        if (series.dataLabel() != null) writeDataLabels(xml, series.dataLabel());

        if (series.categories() != null) {
            String catTag = chartType == ChartType.SCATTER ? "xVal" : "cat";
            writeDataSourceElement(xml, catTag, series.categories(), sheetName);
        }

        String valTag = chartType == ChartType.SCATTER ? "yVal" : "val";
        writeDataSourceElement(xml, valTag, series.values(), sheetName);

        if (series.smooth()) cEmptyVal(xml, "smooth", "1");
        if (series.explosion() > 0) cEmptyVal(xml, "explosion", String.valueOf(series.explosion()));

        xml.endElement(); // c:ser
    }

    // --- Data sources ---

    private static void writeDataSourceElement(XmlWriter xml, String localName,
                                                ChartDataSource source, String sheetName) throws IOException {
        cStart(xml, localName);
        boolean isNumeric = localName.equals("val") || localName.equals("yVal") || localName.equals("xVal");

        switch (source) {
            case ChartDataSource.CellReference ref -> {
                String qualifiedRef = ref.qualify(sheetName).reference();
                String refTag = isNumeric ? "numRef" : "strRef";
                cStart(xml, refTag);
                cStart(xml, "f"); xml.text(qualifiedRef); xml.endElement();
                xml.endElement();
            }
            case ChartDataSource.NumberArray arr -> {
                cStart(xml, "numLit");
                cEmptyVal(xml, "ptCount", String.valueOf(arr.values().size()));
                for (int i = 0; i < arr.values().size(); i++) {
                    cStart(xml, "pt"); xml.attribute("idx", String.valueOf(i));
                    cStart(xml, "v"); xml.text(String.valueOf(arr.values().get(i))); xml.endElement();
                    xml.endElement();
                }
                xml.endElement();
            }
            case ChartDataSource.StringArray arr -> {
                cStart(xml, "strLit");
                cEmptyVal(xml, "ptCount", String.valueOf(arr.values().size()));
                for (int i = 0; i < arr.values().size(); i++) {
                    cStart(xml, "pt"); xml.attribute("idx", String.valueOf(i));
                    cStart(xml, "v"); xml.text(arr.values().get(i)); xml.endElement();
                    xml.endElement();
                }
                xml.endElement();
            }
        }
        xml.endElement();
    }

    // --- Axes ---

    private static void writeAxes(XmlWriter xml, Chart chart) throws IOException {
        List<ChartAxis> axes = chart.axes();

        if (axes.isEmpty()) {
            if (chart.type() == ChartType.SCATTER) {
                writeValueAxis(xml, ValueAxis.of(0, 1));
                writeValueAxis(xml, ValueAxis.of(1, 0));
            } else {
                writeCategoryAxis(xml, CategoryAxis.of(0, 1));
                writeValueAxis(xml, ValueAxis.of(1, 0));
            }
        } else {
            for (ChartAxis axis : axes) {
                switch (axis) {
                    case CategoryAxis cat -> writeCategoryAxis(xml, cat);
                    case ValueAxis val -> writeValueAxis(xml, val);
                    case DateAxis date -> writeDateAxis(xml, date);
                    case SeriesAxis ser -> writeSeriesAxis(xml, ser);
                }
            }
        }
    }

    private static void writeCategoryAxis(XmlWriter xml, CategoryAxis axis) throws IOException {
        cStart(xml, "catAx");
        writeAxisCommon(xml, axis);
        cEmptyVal(xml, "crosses", "autoZero");
        cEmptyVal(xml, "auto", "false");
        xml.endElement();
    }

    private static void writeValueAxis(XmlWriter xml, ValueAxis axis) throws IOException {
        cStart(xml, "valAx");
        writeAxisCommon(xml, axis);
        cEmptyVal(xml, "crosses", xmlValue(axis.crosses()));
        cEmptyVal(xml, "crossBetween", xmlValue(axis.crossBetween()));
        if (axis.majorUnit() != null) cEmptyVal(xml, "majorUnit", String.valueOf(axis.majorUnit()));
        if (axis.minorUnit() != null) cEmptyVal(xml, "minorUnit", String.valueOf(axis.minorUnit()));
        xml.endElement();
    }

    private static void writeDateAxis(XmlWriter xml, DateAxis axis) throws IOException {
        cStart(xml, "dateAx");
        writeAxisCommon(xml, axis);
        xml.endElement();
    }

    private static void writeSeriesAxis(XmlWriter xml, SeriesAxis axis) throws IOException {
        cStart(xml, "serAx");
        writeAxisCommon(xml, axis);
        xml.endElement();
    }

    private static void writeAxisCommon(XmlWriter xml, ChartAxis axis) throws IOException {
        cEmptyVal(xml, "axId", String.valueOf(axis.id()));

        cStart(xml, "scaling");
        cEmptyVal(xml, "orientation", xmlValue(axis.orientation()));
        if (axis instanceof ValueAxis va) {
            if (va.minimum() != null) cEmptyVal(xml, "min", String.valueOf(va.minimum()));
            if (va.maximum() != null) cEmptyVal(xml, "max", String.valueOf(va.maximum()));
            if (va.logBase() != null) cEmptyVal(xml, "logBase", String.valueOf(va.logBase()));
        }
        xml.endElement(); // c:scaling

        cEmptyVal(xml, "delete", axis.visible() ? "false" : "true");
        cEmptyVal(xml, "axPos", xmlValue(axis.position()));

        if (axis.title() != null) writeAxisTitle(xml, axis.title());

        if (axis instanceof ValueAxis va && va.numberFormat() != null) {
            cEmpty(xml, "numFmt");
            xml.attribute("formatCode", va.numberFormat());
            xml.attribute("sourceLinked", "0");
        }

        cEmptyVal(xml, "majorTickMark", xmlValue(axis.majorTickMark()));
        cEmptyVal(xml, "minorTickMark", xmlValue(axis.minorTickMark()));
        cEmptyVal(xml, "tickLblPos", "nextTo");
        cEmptyVal(xml, "crossAx", String.valueOf(axis.crossAxisId()));
    }

    private static void writeAxisTitle(XmlWriter xml, String titleText) throws IOException {
        cStart(xml, "title");
        cStart(xml, "tx");
        cStart(xml, "rich");
        aEmpty(xml, "bodyPr");
        aEmpty(xml, "lstStyle");
        aStart(xml, "p");
        aStart(xml, "pPr");
        aEmpty(xml, "defRPr");
        xml.endElement(); // a:pPr
        aStart(xml, "r");
        aEmpty(xml, "rPr");
        xml.attribute("lang", "en-US");
        aStart(xml, "t"); xml.text(titleText); xml.endElement();
        xml.endElement(); // a:r
        xml.endElement(); // a:p
        xml.endElement(); // c:rich
        xml.endElement(); // c:tx
        cEmptyVal(xml, "overlay", "0");
        xml.endElement(); // c:title
    }

    // --- Legend ---

    private static void writeLegend(XmlWriter xml, ChartLegend legend) throws IOException {
        cStart(xml, "legend");
        cEmptyVal(xml, "legendPos", xmlValue(legend.position()));
        cEmptyVal(xml, "overlay", legend.overlay() ? "true" : "false");
        xml.endElement();
    }

    // --- Style: fill, color, line ---

    private static void writeFill(XmlWriter xml, ChartFill fill) throws IOException {
        switch (fill) {
            case ChartFill.Solid solid -> {
                aStart(xml, "solidFill"); writeColor(xml, solid.color()); xml.endElement();
            }
            case ChartFill.Gradient grad -> {
                aStart(xml, "gradFill");
                aStart(xml, "gsLst");
                for (GradientStop stop : grad.stops()) {
                    aStart(xml, "gs");
                    xml.attribute("pos", String.valueOf((int) (stop.position() * 100000)));
                    writeColor(xml, stop.color());
                    xml.endElement();
                }
                xml.endElement(); // a:gsLst
                aEmpty(xml, "lin");
                xml.attribute("ang", String.valueOf((int) (grad.angle() * 60000)));
                xml.attribute("scaled", "1");
                xml.endElement();
            }
            case ChartFill.Pattern pat -> {
                aStart(xml, "pattFill");
                xml.attribute("prst", xmlValue(pat.type()));
                aStart(xml, "fgClr"); writeColor(xml, pat.foreground()); xml.endElement();
                aStart(xml, "bgClr"); writeColor(xml, pat.background()); xml.endElement();
                xml.endElement();
            }
            case ChartFill.Picture _ -> aEmpty(xml, "noFill");
            case ChartFill.None _ -> aEmpty(xml, "noFill");
        }
    }

    private static void writeColor(XmlWriter xml, ChartColor color) throws IOException {
        switch (color) {
            case ChartColor.Rgb rgb -> { aEmpty(xml, "srgbClr"); xml.attribute("val", rgb.hex()); }
            case ChartColor.Preset p -> { aEmpty(xml, "prstClr"); xml.attribute("val", xmlValue(p.color())); }
            case ChartColor.Theme t -> { aEmpty(xml, "schemeClr"); xml.attribute("val", xmlValue(t.color())); }
        }
    }

    private static void writeLine(XmlWriter xml, ChartLine line) throws IOException {
        aStart(xml, "ln");
        xml.attribute("w", String.valueOf((int) (line.width() * 12700)));
        if (line.color() != null) {
            aStart(xml, "solidFill"); writeColor(xml, line.color()); xml.endElement();
        }
        aEmpty(xml, "prstDash");
        xml.attribute("val", xmlValue(line.dash()));
        xml.endElement(); // a:ln
    }

    private static void writeMarker(XmlWriter xml, ChartMarker marker) throws IOException {
        cStart(xml, "marker");
        cEmptyVal(xml, "symbol", xmlValue(marker.style()));
        cEmptyVal(xml, "size", String.valueOf(marker.size()));
        if (marker.fill() != null) {
            cStart(xml, "spPr");
            writeFill(xml, marker.fill());
            xml.endElement();
        }
        xml.endElement();
    }

    private static void writeDataLabels(XmlWriter xml, ChartDataLabel label) throws IOException {
        cStart(xml, "dLbls");
        cEmptyVal(xml, "showLegendKey", "0");
        cEmptyVal(xml, "showVal", label.showValue() ? "1" : "0");
        cEmptyVal(xml, "showCatName", label.showCategory() ? "1" : "0");
        cEmptyVal(xml, "showSerName", label.showSeriesName() ? "1" : "0");
        cEmptyVal(xml, "showPercent", label.showPercent() ? "1" : "0");
        cEmptyVal(xml, "showBubbleSize", "0");
        if (label.separator() != null) {
            cStart(xml, "separator"); xml.text(label.separator()); xml.endElement();
        }
        if (label.showLeaderLines()) cEmptyVal(xml, "showLeaderLines", "1");
        xml.endElement();
    }

    private static void writeRunProperties(XmlWriter xml, ChartFont font) throws IOException {
        aStart(xml, "rPr");
        xml.attribute("lang", "en-US");
        xml.attribute("sz", String.valueOf((int) (font.size() * 100)));
        if (font.bold()) xml.attribute("b", "1");
        if (font.italic()) xml.attribute("i", "1");
        if (font.color() != null) {
            aStart(xml, "solidFill"); writeColor(xml, font.color()); xml.endElement();
        }
        if (font.name() != null) {
            aEmpty(xml, "latin");
            xml.attribute("typeface", font.name());
        }
        xml.endElement();
    }

    private static void writePrintSettings(XmlWriter xml) throws IOException {
        cStart(xml, "printSettings");
        cEmpty(xml, "headerFooter");
        cEmpty(xml, "pageMargins");
        xml.attribute("b", "0.75");
        xml.attribute("l", "0.7");
        xml.attribute("r", "0.7");
        xml.attribute("t", "0.75");
        xml.attribute("header", "0.3");
        xml.attribute("footer", "0.3");
        cEmpty(xml, "pageSetup");
        xml.endElement();
    }

    // --- Namespace helpers ---

    /** Starts a c: element */
    private static void cStart(XmlWriter xml, String localName) throws IOException {
        xml.startElement("c", localName, NS_C);
    }

    /** Writes an empty c: element */
    private static void cEmpty(XmlWriter xml, String localName) throws IOException {
        xml.emptyElement("c", localName, NS_C);
    }

    /** Writes {@code <c:name val="value"/>} */
    private static void cEmptyVal(XmlWriter xml, String localName, String value) throws IOException {
        xml.emptyElement("c", localName, NS_C);
        xml.attribute("val", value);
    }

    /** Starts an a: element */
    private static void aStart(XmlWriter xml, String localName) throws IOException {
        xml.startElement("a", localName, NS_A);
    }

    /** Writes an empty a: element */
    private static void aEmpty(XmlWriter xml, String localName) throws IOException {
        xml.emptyElement("a", localName, NS_A);
    }

    private ChartXmlWriter() {}
}
