package com.beingidly.litexl.format;

import com.beingidly.litexl.CellRange;
import org.jspecify.annotations.Nullable;

import java.util.List;

/**
 * Represents an AutoFilter on a worksheet.
 */
public record AutoFilter(CellRange range, List<FilterColumn> columns) {

    /**
     * Represents a filter on a single column.
     */
    public record FilterColumn(
        int columnIndex,
        List<String> values,
        @Nullable CustomFilter custom
    ) {
        /**
         * Creates a value-based filter.
         */
        public static FilterColumn values(int columnIndex, List<String> values) {
            return new FilterColumn(columnIndex, values, null);
        }

        /**
         * Creates a custom filter.
         */
        public static FilterColumn custom(int columnIndex, CustomFilter filter) {
            return new FilterColumn(columnIndex, List.of(), filter);
        }
    }

    /**
     * Represents a custom filter condition.
     */
    public record CustomFilter(
        Operator op1,
        String val1,
        @Nullable Operator op2,
        @Nullable String val2,
        boolean and
    ) {
        public enum Operator {
            EQUAL,
            NOT_EQUAL,
            GREATER_THAN,
            GREATER_THAN_OR_EQUAL,
            LESS_THAN,
            LESS_THAN_OR_EQUAL
        }

        /**
         * Creates a single condition filter.
         */
        public static CustomFilter single(Operator op, String value) {
            return new CustomFilter(op, value, null, null, true);
        }

        /**
         * Creates an AND filter with two conditions.
         */
        public static CustomFilter and(Operator op1, String val1, Operator op2, String val2) {
            return new CustomFilter(op1, val1, op2, val2, true);
        }

        /**
         * Creates an OR filter with two conditions.
         */
        public static CustomFilter or(Operator op1, String val1, Operator op2, String val2) {
            return new CustomFilter(op1, val1, op2, val2, false);
        }
    }
}
