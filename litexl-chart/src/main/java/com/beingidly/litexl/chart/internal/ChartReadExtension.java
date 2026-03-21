package com.beingidly.litexl.chart.internal;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.XmlReader;
import com.beingidly.litexl.chart.Chart;
import com.beingidly.litexl.chart.ChartPosition;
import com.beingidly.litexl.chart.Charts;
import com.beingidly.litexl.spi.ReadContext;
import com.beingidly.litexl.spi.ReadExtension;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * ReadExtension implementation for charts.
 * Discovered via ServiceLoader.
 */
public final class ChartReadExtension implements ReadExtension {

    @Override
    public void read(ReadContext ctx, Workbook workbook) throws IOException {
        for (int i = 0; i < workbook.sheetCount(); i++) {
            Sheet sheet = workbook.getSheet(i);
            if (sheet == null) continue;

            int sheetNum = i + 1;
            String sheetRelsPath = "xl/worksheets/_rels/sheet" + sheetNum + ".xml.rels";

            if (!ctx.hasEntry(sheetRelsPath)) continue;

            // Parse sheet relationships to find drawing reference
            Map<String, String> rels = parseRelationships(ctx, sheetRelsPath);
            String drawingTarget = findDrawingTarget(rels);
            if (drawingTarget == null) continue;

            // Resolve drawing path relative to worksheet
            String drawingPath = resolveRelativePath("xl/worksheets/", drawingTarget);

            if (!ctx.hasEntry(drawingPath)) continue;

            // Parse drawing XML to find chart references and positions
            List<DrawingXmlReader.ChartRef> chartRefs;
            try (InputStream drawingIs = ctx.openEntry(drawingPath)) {
                if (drawingIs == null) continue;
                chartRefs = DrawingXmlReader.read(drawingIs);
            }

            if (chartRefs.isEmpty()) continue;

            // Parse drawing relationships to resolve chart file paths
            String drawingDir = drawingPath.substring(0, drawingPath.lastIndexOf('/') + 1);
            String drawingRelsPath = drawingDir + "_rels/"
                + drawingPath.substring(drawingPath.lastIndexOf('/') + 1) + ".rels";

            Map<String, String> drawingRels = ctx.hasEntry(drawingRelsPath)
                ? parseRelationships(ctx, drawingRelsPath)
                : Map.of();

            // Read each chart
            for (DrawingXmlReader.ChartRef ref : chartRefs) {
                String chartTarget = drawingRels.get(ref.rId());
                if (chartTarget == null) continue;

                String chartPath = resolveRelativePath(drawingDir, chartTarget);
                if (!ctx.hasEntry(chartPath)) continue;

                try (InputStream chartIs = ctx.openEntry(chartPath)) {
                    if (chartIs == null) continue;
                    Chart chart = ChartXmlReader.read(chartIs, ref.position());
                    Charts.add(sheet, chart);
                }
            }
        }
    }

    private Map<String, String> parseRelationships(ReadContext ctx, String path) throws IOException {
        Map<String, String> rels = new HashMap<>();
        try (InputStream is = ctx.openEntry(path)) {
            if (is == null) return rels;
            XmlReader xml = new XmlReader(is);
            while (xml.hasNext()) {
                XmlReader.Event event = xml.next();
                if (event == XmlReader.Event.START_ELEMENT
                    && "Relationship".equals(xml.getLocalName())) {
                    String id = xml.getAttributeValue("Id");
                    String target = xml.getAttributeValue("Target");
                    if (id != null && target != null) {
                        rels.put(id, target);
                    }
                }
            }
        }
        return rels;
    }

    private String findDrawingTarget(Map<String, String> rels) {
        // Look for a drawing relationship type in the values
        // The rels map is id -> target, but we need to check if any target
        // points to a drawing. In OOXML, drawing targets end with drawingN.xml
        for (String target : rels.values()) {
            if (target.contains("drawing")) {
                return target;
            }
        }
        return null;
    }

    private String resolveRelativePath(String basePath, String relativePath) {
        if (relativePath.startsWith("../")) {
            // Go up one directory
            String parent = basePath;
            String rel = relativePath;
            while (rel.startsWith("../")) {
                rel = rel.substring(3);
                // Remove trailing slash and go up
                if (parent.endsWith("/")) {
                    parent = parent.substring(0, parent.length() - 1);
                }
                int lastSlash = parent.lastIndexOf('/');
                if (lastSlash >= 0) {
                    parent = parent.substring(0, lastSlash + 1);
                }
            }
            return parent + rel;
        }
        return basePath + relativePath;
    }
}
