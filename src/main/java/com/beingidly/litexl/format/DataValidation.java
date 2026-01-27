package com.beingidly.litexl.format;

import com.beingidly.litexl.CellRange;
import org.jspecify.annotations.Nullable;

import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Represents a data validation rule.
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
    public enum Type {
        ANY,
        WHOLE,
        DECIMAL,
        LIST,
        DATE,
        TIME,
        TEXT_LENGTH,
        CUSTOM
    }

    public enum Operator {
        BETWEEN,
        NOT_BETWEEN,
        EQUAL,
        NOT_EQUAL,
        GREATER_THAN,
        LESS_THAN,
        GREATER_THAN_OR_EQUAL,
        LESS_THAN_OR_EQUAL
    }

    /**
     * Creates a list validation with explicit items.
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
