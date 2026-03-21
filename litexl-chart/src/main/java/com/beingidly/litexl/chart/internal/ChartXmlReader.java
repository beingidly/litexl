package com.beingidly.litexl.chart.internal;

import com.beingidly.litexl.XmlReader;
import com.beingidly.litexl.chart.*;
import com.beingidly.litexl.chart.axis.*;
import com.beingidly.litexl.chart.style.*;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Reads chart XML (xl/charts/chartN.xml) to build Chart objects.
 */
final class ChartXmlReader {

    static Chart read(InputStream is, ChartPosition position) throws IOException {
        try (XmlReader xml = new XmlReader(is)) {
            ChartType type = null;
            String titleText = null;
            List<ChartSeries> seriesList = new ArrayList<>();
            LegendPosition legendPos = null;
            Grouping grouping = null;
            BarDirection barDirection = null;
            ScatterStyle scatterStyle = null;

            boolean inTitle = false;
            boolean inSeries = false;
            boolean inLegend = false;
            String seriesName = null;
            String catRef = null;
            String valRef = null;

            while (xml.hasNext()) {
                XmlReader.Event event = xml.next();
                if (event == XmlReader.Event.START_ELEMENT) {
                    String name = xml.getLocalName();

                    // Detect chart type
                    ChartType detected = detectChartType(name);
                    if (detected != null) {
                        type = detected;
                    }

                    switch (name) {
                        case "title" -> { if (type == null) inTitle = true; }
                        case "ser" -> { inSeries = true; seriesName = null; catRef = null; valRef = null; }
                        case "legend" -> inLegend = true;
                        case "grouping" -> {
                            String val = xml.getAttributeValue("val");
                            if (val != null) grouping = parseGrouping(val);
                        }
                        case "barDir" -> {
                            String val = xml.getAttributeValue("val");
                            if (val != null) barDirection = "bar".equals(val) ? BarDirection.BAR : BarDirection.COLUMN;
                        }
                        case "scatterStyle" -> {
                            String val = xml.getAttributeValue("val");
                            if (val != null) scatterStyle = parseScatterStyle(val);
                        }
                        case "legendPos" -> {
                            if (inLegend) {
                                String val = xml.getAttributeValue("val");
                                if (val != null) legendPos = parseLegendPosition(val);
                            }
                        }
                        case "f" -> {
                            // formula element - read the cell reference
                            if (inSeries) {
                                String text = xml.getElementText();
                                // Determine if this is cat or val by context
                                // The last parent element determines this
                                // We'll use a simple heuristic: first f is cat, second is val
                                if (catRef == null && valRef == null) {
                                    catRef = text;
                                } else if (valRef == null) {
                                    valRef = text;
                                }
                            }
                        }
                        case "v" -> {
                            if (inSeries && seriesName == null && inTitle) {
                                // This could be series name
                            }
                        }
                        case "t" -> {
                            if (inTitle && !inSeries) {
                                titleText = xml.getElementText();
                            }
                        }
                    }
                } else if (event == XmlReader.Event.END_ELEMENT) {
                    String name = xml.getLocalName();
                    switch (name) {
                        case "title" -> inTitle = false;
                        case "legend" -> inLegend = false;
                        case "ser" -> {
                            inSeries = false;
                            // Build series from collected data
                            ChartDataSource catSource = catRef != null ? ChartDataSource.ofRange(catRef) : null;
                            ChartDataSource valSource = valRef != null
                                ? ChartDataSource.ofRange(valRef)
                                : ChartDataSource.ofNumbers(0);
                            seriesList.add(new ChartSeries(
                                seriesName, catSource, valSource,
                                null, null, null, null, null, false, 0));
                            seriesName = null;
                            catRef = null;
                            valRef = null;
                        }
                    }
                }
            }

            if (type == null) {
                type = ChartType.BAR; // fallback
            }
            if (seriesList.isEmpty()) {
                seriesList.add(ChartSeries.of("A1:A1")); // placeholder
            }

            ChartTitle title = titleText != null ? ChartTitle.of(titleText) : null;
            ChartLegend legend = legendPos != null ? ChartLegend.of(legendPos) : null;
            ChartPlotConfig plotConfig = new ChartPlotConfig(
                grouping, barDirection, scatterStyle, null, null,
                DisplayBlanks.GAP, true);

            return new Chart(type, title, position, seriesList, legend, plotConfig, List.of());
        }
    }

    private static ChartType detectChartType(String localName) {
        return switch (localName) {
            case "barChart" -> ChartType.BAR; // direction determined by barDir
            case "lineChart" -> ChartType.LINE;
            case "pieChart" -> ChartType.PIE;
            case "scatterChart" -> ChartType.SCATTER;
            case "areaChart" -> ChartType.AREA;
            case "radarChart" -> ChartType.RADAR;
            case "doughnutChart" -> ChartType.DOUGHNUT;
            case "surfaceChart" -> ChartType.SURFACE;
            case "bar3DChart" -> ChartType.BAR_3D;
            case "line3DChart" -> ChartType.LINE_3D;
            case "area3DChart" -> ChartType.AREA_3D;
            case "pie3DChart" -> ChartType.PIE_3D;
            default -> null;
        };
    }

    private static Grouping parseGrouping(String val) {
        return switch (val) {
            case "standard" -> Grouping.STANDARD;
            case "clustered" -> Grouping.CLUSTERED;
            case "stacked" -> Grouping.STACKED;
            case "percentStacked" -> Grouping.PERCENT_STACKED;
            default -> null;
        };
    }

    private static ScatterStyle parseScatterStyle(String val) {
        return switch (val) {
            case "marker" -> ScatterStyle.MARKER;
            case "line" -> ScatterStyle.LINE;
            case "lineMarker" -> ScatterStyle.LINE_MARKER;
            case "smooth" -> ScatterStyle.SMOOTH;
            case "smoothMarker" -> ScatterStyle.SMOOTH_MARKER;
            default -> null;
        };
    }

    private static LegendPosition parseLegendPosition(String val) {
        return switch (val) {
            case "b" -> LegendPosition.BOTTOM;
            case "t" -> LegendPosition.TOP;
            case "l" -> LegendPosition.LEFT;
            case "r" -> LegendPosition.RIGHT;
            default -> LegendPosition.BOTTOM;
        };
    }

    private ChartXmlReader() {}
}
