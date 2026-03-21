package com.beingidly.litexl.format;

import com.beingidly.litexl.CellRange;
import org.jspecify.annotations.Nullable;

import java.util.List;

/**
 * Represents an AutoFilter on a worksheet.
 *
 * @param range the cell range the filter applies to
 * @param columns the filter column definitions
 */
public record AutoFilter(CellRange range, List<FilterColumn> columns) {

    /**
     * Represents a filter on a single column.
     *
     * @param columnIndex the zero-based column index
     * @param values the filter values for value-based filtering
     * @param custom the custom filter condition, or null
     */
    public record FilterColumn(
        int columnIndex,
        List<String> values,
        @Nullable CustomFilter custom
    ) {
        /**
         * Creates a value-based filter.
         *
         * @param columnIndex the zero-based column index
         * @param values the filter values
         * @return a new value-based filter column
         */
        public static FilterColumn values(int columnIndex, List<String> values) {
            return new FilterColumn(columnIndex, values, null);
        }

        /**
         * Creates a custom filter.
         *
         * @param columnIndex the zero-based column index
         * @param filter the custom filter condition
         * @return a new custom filter column
         */
        public static FilterColumn custom(int columnIndex, CustomFilter filter) {
            return new FilterColumn(columnIndex, List.of(), filter);
        }
    }

    /**
     * Represents a custom filter condition.
     *
     * @param op1 the first operator
     * @param val1 the first comparison value
     * @param op2 the second operator, or null for single-condition filters
     * @param val2 the second comparison value, or null for single-condition filters
     * @param and true to combine with AND, false for OR
     */
    public record CustomFilter(
        Operator op1,
        String val1,
        @Nullable Operator op2,
        @Nullable String val2,
        boolean and
    ) {
        /** Comparison operators for custom filters. */
        public enum Operator {
            /** Equal to. */
            EQUAL,
            /** Not equal to. */
            NOT_EQUAL,
            /** Greater than. */
            GREATER_THAN,
            /** Greater than or equal to. */
            GREATER_THAN_OR_EQUAL,
            /** Less than. */
            LESS_THAN,
            /** Less than or equal to. */
            LESS_THAN_OR_EQUAL
        }

        /**
         * Creates a single condition filter.
         *
         * @param op the comparison operator
         * @param value the comparison value
         * @return a new single-condition custom filter
         */
        public static CustomFilter single(Operator op, String value) {
            return new CustomFilter(op, value, null, null, true);
        }

        /**
         * Creates an AND filter with two conditions.
         *
         * @param op1 the first operator
         * @param val1 the first comparison value
         * @param op2 the second operator
         * @param val2 the second comparison value
         * @return a new AND custom filter
         */
        public static CustomFilter and(Operator op1, String val1, Operator op2, String val2) {
            return new CustomFilter(op1, val1, op2, val2, true);
        }

        /**
         * Creates an OR filter with two conditions.
         *
         * @param op1 the first operator
         * @param val1 the first comparison value
         * @param op2 the second operator
         * @param val2 the second comparison value
         * @return a new OR custom filter
         */
        public static CustomFilter or(Operator op1, String val1, Operator op2, String val2) {
            return new CustomFilter(op1, val1, op2, val2, false);
        }
    }
}
