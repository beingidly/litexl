package com.beingidly.litexl.format;

import com.beingidly.litexl.CellRange;
import org.jspecify.annotations.Nullable;

import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Represents a data validation rule.
 *
 * @param range the cell range the validation applies to
 * @param type the validation type
 * @param operator the comparison operator, or null
 * @param formula1 the first formula or value, or null
 * @param formula2 the second formula or value, or null
 * @param errorTitle the error dialog title, or null
 * @param errorMessage the error dialog message, or null
 * @param showDropdown whether to show a dropdown list for list validations
 */
public record DataValidation(
    CellRange range,
    Type type,
    @Nullable Operator operator,
    @Nullable String formula1,
    @Nullable String formula2,
    @Nullable String errorTitle,
    @Nullable String errorMessage,
    boolean showDropdown
) {
    /** Data validation types. */
    public enum Type {
        /** Any value allowed. */
        ANY,
        /** Whole number validation. */
        WHOLE,
        /** Decimal number validation. */
        DECIMAL,
        /** List validation. */
        LIST,
        /** Date validation. */
        DATE,
        /** Time validation. */
        TIME,
        /** Text length validation. */
        TEXT_LENGTH,
        /** Custom formula validation. */
        CUSTOM
    }

    /** Comparison operators for data validation. */
    public enum Operator {
        /** Between two values (inclusive). */
        BETWEEN,
        /** Not between two values. */
        NOT_BETWEEN,
        /** Equal to. */
        EQUAL,
        /** Not equal to. */
        NOT_EQUAL,
        /** Greater than. */
        GREATER_THAN,
        /** Less than. */
        LESS_THAN,
        /** Greater than or equal to. */
        GREATER_THAN_OR_EQUAL,
        /** Less than or equal to. */
        LESS_THAN_OR_EQUAL
    }

    /**
     * Creates a list validation with explicit items.
     *
     * @param range the cell range to validate
     * @param items the allowed list items
     * @return a new list data validation
     */
    public static DataValidation list(CellRange range, String... items) {
        String formula = "\"" + Stream.of(items)
            .map(s -> s.replace("\"", "\"\""))
            .collect(Collectors.joining(",")) + "\"";
        return new DataValidation(
            range,
            Type.LIST,
            null,
            formula,
            null,
            null,
            null,
            true
        );
    }

    /**
     * Creates a list validation referencing a cell range.
     *
     * @param range the cell range to validate
     * @param source the cell range containing the list values
     * @return a new list data validation
     */
    public static DataValidation list(CellRange range, CellRange source) {
        return new DataValidation(
            range,
            Type.LIST,
            null,
            source.toAbsoluteRef(),
            null,
            null,
            null,
            true
        );
    }

    /**
     * Creates a whole number validation with min and max values.
     *
     * @param range the cell range to validate
     * @param min the minimum allowed value
     * @param max the maximum allowed value
     * @return a new whole number data validation
     */
    public static DataValidation wholeNumber(CellRange range, int min, int max) {
        return new DataValidation(
            range,
            Type.WHOLE,
            Operator.BETWEEN,
            String.valueOf(min),
            String.valueOf(max),
            "Invalid Input",
            "Please enter a whole number between " + min + " and " + max,
            false
        );
    }

    /**
     * Creates a whole number validation with specified operator and formula values.
     *
     * @param range the cell range to validate
     * @param operator the comparison operator
     * @param formula1 the first formula or value, or null
     * @param formula2 the second formula or value, or null
     * @return a new whole number data validation
     */
    public static DataValidation wholeNumber(CellRange range, Operator operator, @Nullable String formula1, @Nullable String formula2) {
        return new DataValidation(
            range,
            Type.WHOLE,
            operator,
            formula1,
            formula2,
            "Invalid Input",
            "Please enter a valid whole number",
            false
        );
    }

    /**
     * Creates a decimal number validation with min and max values.
     *
     * @param range the cell range to validate
     * @param min the minimum allowed value
     * @param max the maximum allowed value
     * @return a new decimal data validation
     */
    public static DataValidation decimal(CellRange range, double min, double max) {
        return new DataValidation(
            range,
            Type.DECIMAL,
            Operator.BETWEEN,
            String.valueOf(min),
            String.valueOf(max),
            "Invalid Input",
            "Please enter a number between " + min + " and " + max,
            false
        );
    }

    /**
     * Creates a decimal number validation with specified operator and formula values.
     *
     * @param range the cell range to validate
     * @param operator the comparison operator
     * @param formula1 the first formula or value, or null
     * @param formula2 the second formula or value, or null
     * @return a new decimal data validation
     */
    public static DataValidation decimal(CellRange range, Operator operator, @Nullable String formula1, @Nullable String formula2) {
        return new DataValidation(
            range,
            Type.DECIMAL,
            operator,
            formula1,
            formula2,
            "Invalid Input",
            "Please enter a valid decimal number",
            false
        );
    }

    /**
     * Creates a text length validation with min and max lengths.
     *
     * @param range the cell range to validate
     * @param minLength the minimum text length
     * @param maxLength the maximum text length
     * @return a new text length data validation
     */
    public static DataValidation textLength(CellRange range, int minLength, int maxLength) {
        return new DataValidation(
            range,
            Type.TEXT_LENGTH,
            Operator.BETWEEN,
            String.valueOf(minLength),
            String.valueOf(maxLength),
            "Invalid Input",
            "Text length must be between " + minLength + " and " + maxLength + " characters",
            false
        );
    }

    /**
     * Creates a text length validation with specified operator and formula values.
     *
     * @param range the cell range to validate
     * @param operator the comparison operator
     * @param formula1 the first formula or value, or null
     * @param formula2 the second formula or value, or null
     * @return a new text length data validation
     */
    public static DataValidation textLength(CellRange range, Operator operator, @Nullable String formula1, @Nullable String formula2) {
        return new DataValidation(
            range,
            Type.TEXT_LENGTH,
            operator,
            formula1,
            formula2,
            "Invalid Input",
            "Please enter text with valid length",
            false
        );
    }

    /**
     * Creates a custom formula validation with error message.
     *
     * @param range the cell range to validate
     * @param formula the custom validation formula
     * @param errorMessage the error message to display on invalid input
     * @return a new custom data validation
     */
    public static DataValidation custom(CellRange range, String formula, String errorMessage) {
        return new DataValidation(
            range,
            Type.CUSTOM,
            null,
            formula,
            null,
            "Invalid Input",
            errorMessage,
            false
        );
    }

    /**
     * Creates a custom formula validation with default error message.
     *
     * @param range the cell range to validate
     * @param formula the custom validation formula
     * @return a new custom data validation
     */
    public static DataValidation custom(CellRange range, String formula) {
        return new DataValidation(
            range,
            Type.CUSTOM,
            null,
            formula,
            null,
            "Invalid Input",
            "Value does not meet the validation criteria",
            false
        );
    }
}
