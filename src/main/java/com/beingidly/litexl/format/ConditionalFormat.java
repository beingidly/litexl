package com.beingidly.litexl.format;

import com.beingidly.litexl.CellRange;
import org.jspecify.annotations.Nullable;

/**
 * Represents a conditional formatting rule.
 */
public record ConditionalFormat(
    CellRange range,
    Type type,
    Operator operator,
    @Nullable String formula1,
    @Nullable String formula2,
    int styleId
) {
    public enum Type {
        CELL_VALUE,
        EXPRESSION,
        COLOR_SCALE,
        DATA_BAR,
        ICON_SET,
        TOP_BOTTOM,
        ABOVE_AVERAGE,
        DUPLICATE_VALUES,
        UNIQUE_VALUES,
        CONTAINS_TEXT,
        NOT_CONTAINS_TEXT,
        BEGINS_WITH,
        ENDS_WITH,
        CONTAINS_BLANKS,
        CONTAINS_ERRORS
    }

    public enum Operator {
        NONE,
        LESS_THAN,
        LESS_THAN_OR_EQUAL,
        EQUAL,
        NOT_EQUAL,
        GREATER_THAN_OR_EQUAL,
        GREATER_THAN,
        BETWEEN,
        NOT_BETWEEN
    }

    /**
     * Creates a "greater than" conditional format.
     */
    public static ConditionalFormat greaterThan(CellRange range, double value, int styleId) {
        return new ConditionalFormat(
            range,
            Type.CELL_VALUE,
            Operator.GREATER_THAN,
            String.valueOf(value),
            null,
            styleId
        );
    }

    /**
     * Creates a "less than" conditional format.
     */
    public static ConditionalFormat lessThan(CellRange range, double value, int styleId) {
        return new ConditionalFormat(
            range,
            Type.CELL_VALUE,
            Operator.LESS_THAN,
            String.valueOf(value),
            null,
            styleId
        );
    }

    /**
     * Creates a "between" conditional format.
     */
    public static ConditionalFormat between(CellRange range, double min, double max, int styleId) {
        return new ConditionalFormat(
            range,
            Type.CELL_VALUE,
            Operator.BETWEEN,
            String.valueOf(min),
            String.valueOf(max),
            styleId
        );
    }

    /**
     * Creates an expression-based conditional format.
     */
    public static ConditionalFormat expression(CellRange range, String formula, int styleId) {
        return new ConditionalFormat(
            range,
            Type.EXPRESSION,
            Operator.NONE,
            formula,
            null,
            styleId
        );
    }

    /**
     * Creates a "duplicate values" conditional format.
     */
    public static ConditionalFormat duplicateValues(CellRange range, int styleId) {
        return new ConditionalFormat(
            range,
            Type.DUPLICATE_VALUES,
            Operator.NONE,
            null,
            null,
            styleId
        );
    }

    /**
     * Creates a "unique values" conditional format.
     */
    public static ConditionalFormat uniqueValues(CellRange range, int styleId) {
        return new ConditionalFormat(
            range,
            Type.UNIQUE_VALUES,
            Operator.NONE,
            null,
            null,
            styleId
        );
    }
}
