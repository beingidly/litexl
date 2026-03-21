package com.beingidly.litexl.chart;

import java.util.List;

/**
 * Data source for chart series values or categories.
 */
public sealed interface ChartDataSource {

    /**
     * Cell range reference (e.g. "Sheet1!$A$2:$A$13" or "$A$2:$A$13").
     *
     * <p>If the reference does not include a sheet name (no '!'),
     * the owning sheet's name is automatically added at write time.
     */
    record CellReference(String reference) implements ChartDataSource {

        /** Returns true if the reference includes a sheet name. */
        public boolean isQualified() {
            return reference.contains("!");
        }

        /** Returns a qualified reference with the given sheet name if not already qualified. */
        public CellReference qualify(String sheetName) {
            if (isQualified()) {
                return this;
            }
            return new CellReference(escapeSheetName(sheetName) + "!" + reference);
        }

        private static String escapeSheetName(String name) {
            if (name.contains(" ") || name.contains("'")) {
                return "'" + name.replace("'", "''") + "'";
            }
            return name;
        }
    }

    /** Literal numeric values. */
    record NumberArray(List<Double> values) implements ChartDataSource {
        public NumberArray { values = List.copyOf(values); }
    }

    /** Literal string values. */
    record StringArray(List<String> values) implements ChartDataSource {
        public StringArray { values = List.copyOf(values); }
    }

    /** Creates a cell range data source. */
    static ChartDataSource ofRange(String reference) {
        return new CellReference(reference);
    }

    /** Creates a literal number data source. */
    static ChartDataSource ofNumbers(double... values) {
        return new NumberArray(java.util.Arrays.stream(values).boxed().toList());
    }

    /** Creates a literal number data source. */
    static ChartDataSource ofNumbers(List<Double> values) {
        return new NumberArray(values);
    }

    /** Creates a literal string data source. */
    static ChartDataSource ofStrings(String... values) {
        return new StringArray(List.of(values));
    }

    /** Creates a literal string data source. */
    static ChartDataSource ofStrings(List<String> values) {
        return new StringArray(values);
    }
}
