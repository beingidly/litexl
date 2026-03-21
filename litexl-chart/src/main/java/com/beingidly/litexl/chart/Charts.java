package com.beingidly.litexl.chart;

import com.beingidly.litexl.Sheet;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Static API for managing charts on a sheet.
 *
 * <p>Charts are stored in the sheet's extension data and automatically
 * written to the XLSX file when the workbook is saved (if litexl-chart
 * is on the classpath).
 *
 * <pre>{@code
 * Charts.add(sheet, chart);
 * List<Chart> charts = Charts.get(sheet);
 * }</pre>
 */
public final class Charts {

    static final String KEY = "com.beingidly.litexl.chart.CHARTS";

    private Charts() {}

    /**
     * Adds a chart to the sheet.
     *
     * @param sheet the target sheet
     * @param chart the chart to add
     */
    public static void add(Sheet sheet, Chart chart) {
        getOrCreate(sheet).add(chart);
    }

    /**
     * Returns all charts on the sheet (unmodifiable).
     *
     * @param sheet the sheet to query
     * @return an unmodifiable list of charts
     */
    public static List<Chart> get(Sheet sheet) {
        if (!sheet.hasExtensionData(KEY)) {
            return List.of();
        }
        return Collections.unmodifiableList(getOrCreate(sheet));
    }

    /**
     * Removes a chart by index.
     *
     * @param sheet the sheet containing the chart
     * @param index the index of the chart to remove
     */
    public static void remove(Sheet sheet, int index) {
        getOrCreate(sheet).remove(index);
    }

    /**
     * Removes all charts from the sheet.
     *
     * @param sheet the sheet to clear
     */
    public static void clear(Sheet sheet) {
        sheet.removeExtensionData(KEY);
    }

    @SuppressWarnings("unchecked")
    static List<Chart> getOrCreate(Sheet sheet) {
        Object data = sheet.getExtensionData(KEY);
        if (data instanceof List<?> list) {
            return (List<Chart>) list;
        }
        List<Chart> charts = new ArrayList<>();
        sheet.putExtensionData(KEY, charts);
        return charts;
    }
}
