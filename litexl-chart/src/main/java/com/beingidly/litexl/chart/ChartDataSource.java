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
     *
     * @param reference the cell range reference string
     */
    record CellReference(String reference) implements ChartDataSource {

        /**
         * Returns true if the reference includes a sheet name.
         *
         * @return true if the reference is qualified with a sheet name
         */
        public boolean isQualified() {
            return reference.contains("!");
        }

        /**
         * Returns a qualified reference with the given sheet name if not already qualified.
         *
         * @param sheetName the sheet name to prepend
         * @return a qualified cell reference
         */
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

    /**
     * Literal numeric values.
     *
     * @param values the numeric values
     */
    record NumberArray(List<Double> values) implements ChartDataSource {
        /** Compact constructor that defensively copies the values list. */
        public NumberArray { values = List.copyOf(values); }
    }

    /**
     * Literal string values.
     *
     * @param values the string values
     */
    record StringArray(List<String> values) implements ChartDataSource {
        /** Compact constructor that defensively copies the values list. */
        public StringArray { values = List.copyOf(values); }
    }

    /**
     * Creates a cell range data source.
     *
     * @param reference the cell range reference
     * @return a new cell reference data source
     */
    static ChartDataSource ofRange(String reference) {
        return new CellReference(reference);
    }

    /**
     * Creates a literal number data source from varargs.
     *
     * @param values the numeric values
     * @return a new number array data source
     */
    static ChartDataSource ofNumbers(double... values) {
        return new NumberArray(java.util.Arrays.stream(values).boxed().toList());
    }

    /**
     * Creates a literal number data source from a list.
     *
     * @param values the numeric values
     * @return a new number array data source
     */
    static ChartDataSource ofNumbers(List<Double> values) {
        return new NumberArray(values);
    }

    /**
     * Creates a literal string data source from varargs.
     *
     * @param values the string values
     * @return a new string array data source
     */
    static ChartDataSource ofStrings(String... values) {
        return new StringArray(List.of(values));
    }

    /**
     * Creates a literal string data source from a list.
     *
     * @param values the string values
     * @return a new string array data source
     */
    static ChartDataSource ofStrings(List<String> values) {
        return new StringArray(values);
    }
}
