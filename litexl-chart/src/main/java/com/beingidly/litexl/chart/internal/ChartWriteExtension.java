package com.beingidly.litexl.chart.internal;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.XmlWriter;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.style.ChartFill;
import com.beingidly.litexl.spi.*;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * WriteExtension implementation for charts.
 * Discovered via ServiceLoader.
 */
public final class ChartWriteExtension implements WriteExtension {

    /** Creates a new chart write extension. */
    public ChartWriteExtension() {}

    @Override
    public void contributeContentTypes(ContentTypeRegistry registry, List<Sheet> sheets) throws IOException {
        int chartNum = 0;
        int drawingNum = 0;
        boolean hasImages = false;

        for (int i = 0; i < sheets.size(); i++) {
            List<Chart> charts = Charts.get(sheets.get(i));
            if (!charts.isEmpty()) {
                drawingNum++;
                registry.addOverride("/xl/drawings/drawing" + drawingNum + ".xml",
                    "application/vnd.openxmlformats-officedocument.drawing+xml");

                for (Chart chart : charts) {
                    chartNum++;
                    registry.addOverride("/xl/charts/chart" + chartNum + ".xml",
                        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");

                    if (!hasImages) {
                        hasImages = hasImageFills(chart);
                    }
                }
            }
        }

        if (hasImages) {
            registry.addDefault("png", "image/png");
            registry.addDefault("jpeg", "image/jpeg");
            registry.addDefault("gif", "image/gif");
        }
    }

    @Override
    public void writeSheetElements(XmlWriter xml, Sheet sheet, int sheetNum) throws IOException {
        List<Chart> charts = Charts.get(sheet);
        if (charts.isEmpty()) return;

        // Write <drawing r:id="rId1"/>
        xml.emptyElement("drawing");
        xml.attribute("r:id", "rId1");
    }

    @Override
    public List<Relationship> sheetRelationships(Sheet sheet, int sheetNum) {
        List<Chart> charts = Charts.get(sheet);
        if (charts.isEmpty()) return List.of();

        return List.of(new Relationship("rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            "../drawings/drawing" + sheetNum + ".xml"));
    }

    @Override
    public void writeEntries(WriteContext ctx, List<Sheet> sheets) throws IOException {
        int globalChartNum = 0;
        int globalMediaNum = 0;

        for (int i = 0; i < sheets.size(); i++) {
            Sheet sheet = sheets.get(i);
            List<Chart> charts = Charts.get(sheet);
            if (charts.isEmpty()) continue;

            int sheetNum = i + 1;
            int firstChartInDrawing = globalChartNum + 1;
            List<Integer> chartNums = new ArrayList<>();

            // Write each chart XML
            for (Chart chart : charts) {
                globalChartNum++;
                chartNums.add(globalChartNum);

                // Write media files for picture fills
                List<MediaWriter.MediaEntry> media = MediaWriter.collectMedia(chart, globalMediaNum);
                for (MediaWriter.MediaEntry entry : media) {
                    globalMediaNum++;
                    try (OutputStream os = ctx.newEntry(entry.path())) {
                        os.write(entry.data());
                    }
                }

                try (OutputStream os = ctx.newEntry("xl/charts/chart" + globalChartNum + ".xml");
                     XmlWriter xml = new XmlWriter(os)) {
                    ChartXmlWriter.write(xml, chart, sheet);
                }
            }

            // Write drawing XML
            try (OutputStream os = ctx.newEntry("xl/drawings/drawing" + sheetNum + ".xml");
                 XmlWriter xml = new XmlWriter(os)) {
                DrawingXmlWriter.write(xml, charts, firstChartInDrawing);
            }

            // Write drawing rels
            try (OutputStream os = ctx.newEntry("xl/drawings/_rels/drawing" + sheetNum + ".xml.rels");
                 XmlWriter xml = new XmlWriter(os)) {
                xml.startDocument();
                xml.startElement("Relationships");
                xml.attribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
                for (int j = 0; j < chartNums.size(); j++) {
                    xml.emptyElement("Relationship");
                    xml.attribute("Id", "rId" + (j + 1));
                    xml.attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");
                    xml.attribute("Target", "../charts/chart" + chartNums.get(j) + ".xml");
                }
                xml.endElement();
                xml.endDocument();
            }
        }
    }

    private boolean hasImageFills(Chart chart) {
        for (ChartSeries s : chart.series()) {
            if (s.fill() instanceof ChartFill.Picture) return true;
        }
        return false;
    }
}
