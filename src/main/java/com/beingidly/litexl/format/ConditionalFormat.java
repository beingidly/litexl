package com.beingidly.litexl.format;

import com.beingidly.litexl.CellRange;
import org.jspecify.annotations.Nullable;

/**
 * Represents a conditional formatting rule.
 *
 * @param range the cell range the rule applies to
 * @param type the type of conditional format
 * @param operator the comparison operator
 * @param formula1 the first formula or value, or null
 * @param formula2 the second formula or value, or null
 * @param styleId the style index to apply when the condition is met
 */
public record ConditionalFormat(
    CellRange range,
    Type type,
    Operator operator,
    @Nullable String formula1,
    @Nullable String formula2,
    int styleId
) {
    /** Types of conditional formatting rules. */
    public enum Type {
        /** Compares cell values. */
        CELL_VALUE,
        /** Evaluates a formula expression. */
        EXPRESSION,
        /** Applies a color scale. */
        COLOR_SCALE,
        /** Applies a data bar. */
        DATA_BAR,
        /** Applies an icon set. */
        ICON_SET,
        /** Highlights top or bottom ranked values. */
        TOP_BOTTOM,
        /** Highlights values above or below average. */
        ABOVE_AVERAGE,
        /** Highlights duplicate values. */
        DUPLICATE_VALUES,
        /** Highlights unique values. */
        UNIQUE_VALUES,
        /** Highlights cells containing specific text. */
        CONTAINS_TEXT,
        /** Highlights cells not containing specific text. */
        NOT_CONTAINS_TEXT,
        /** Highlights cells beginning with specific text. */
        BEGINS_WITH,
        /** Highlights cells ending with specific text. */
        ENDS_WITH,
        /** Highlights blank cells. */
        CONTAINS_BLANKS,
        /** Highlights cells containing errors. */
        CONTAINS_ERRORS
    }

    /** Comparison operators for conditional formatting. */
    public enum Operator {
        /** No operator. */
        NONE,
        /** Less than. */
        LESS_THAN,
        /** Less than or equal to. */
        LESS_THAN_OR_EQUAL,
        /** Equal to. */
        EQUAL,
        /** Not equal to. */
        NOT_EQUAL,
        /** Greater than or equal to. */
        GREATER_THAN_OR_EQUAL,
        /** Greater than. */
        GREATER_THAN,
        /** Between two values (inclusive). */
        BETWEEN,
        /** Not between two values. */
        NOT_BETWEEN
    }

    /**
     * Creates a "greater than" conditional format.
     *
     * @param range the cell range to apply the rule to
     * @param value the threshold value
     * @param styleId the style index to apply when the condition is met
     * @return a new conditional format rule
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
     *
     * @param range the cell range to apply the rule to
     * @param value the threshold value
     * @param styleId the style index to apply when the condition is met
     * @return a new conditional format rule
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
     *
     * @param range the cell range to apply the rule to
     * @param min the minimum value (inclusive)
     * @param max the maximum value (inclusive)
     * @param styleId the style index to apply when the condition is met
     * @return a new conditional format rule
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
     *
     * @param range the cell range to apply the rule to
     * @param formula the formula expression
     * @param styleId the style index to apply when the condition is met
     * @return a new conditional format rule
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
     *
     * @param range the cell range to apply the rule to
     * @param styleId the style index to apply when the condition is met
     * @return a new conditional format rule
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
     *
     * @param range the cell range to apply the rule to
     * @param styleId the style index to apply when the condition is met
     * @return a new conditional format rule
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
