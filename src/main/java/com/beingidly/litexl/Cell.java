package com.beingidly.litexl;

import org.jspecify.annotations.Nullable;

import java.time.LocalDateTime;

/**
 * Represents a cell in a worksheet.
 *
 * <p>This class is <b>not thread-safe</b>.
 * External synchronization is required for concurrent access.
 */
public final class Cell {

    private final int column;
    private CellValue value;
    private int styleId;

    Cell(int column) {
        this.column = column;
        this.value = new CellValue.Empty();
        this.styleId = 0;
    }

    /**
     * Returns the column index (0-based).
     */
    public int column() {
        return column;
    }

    /**
     * Returns the cell type.
     */
    public CellType type() {
        return value.type();
    }

    /**
     * Returns the raw cell value.
     */
    public CellValue value() {
        return value;
    }

    /**
     * Returns the string value, or null if not a string cell.
     */
    public @Nullable String string() {
        return value instanceof CellValue.Text(String value1) ? value1 : null;
    }

    /**
     * Returns the numeric value, or 0 if not a number cell.
     */
    public double number() {
        return value instanceof CellValue.Number(double value1) ? value1 : 0.0;
    }

    /**
     * Returns the boolean value, or false if not a boolean cell.
     */
    public boolean bool() {
        return value instanceof CellValue.Bool(boolean value1) && value1;
    }

    /**
     * Returns the date value, or null if not a date cell.
     */
    public @Nullable LocalDateTime date() {
        return value instanceof CellValue.Date(LocalDateTime value1) ? value1 : null;
    }

    /**
     * Returns the formula expression, or null if not a formula cell.
     */
    public @Nullable String formula() {
        return value instanceof CellValue.Formula f ? f.expression() : null;
    }

    /**
     * Returns the error code, or null if not an error cell.
     */
    public @Nullable String error() {
        return value instanceof CellValue.Error(String code) ? code : null;
    }

    /**
     * Sets the cell value to a string.
     */
    public Cell set(@Nullable String value) {
        this.value = value == null ? new CellValue.Empty() : new CellValue.Text(value);
        return this;
    }

    /**
     * Sets the cell value to a number.
     */
    public Cell set(double value) {
        this.value = new CellValue.Number(value);
        return this;
    }

    /**
     * Sets the cell value to a boolean.
     */
    public Cell set(boolean value) {
        this.value = new CellValue.Bool(value);
        return this;
    }

    /**
     * Sets the cell value to a date.
     */
    public Cell set(@Nullable LocalDateTime value) {
        this.value = value == null ? new CellValue.Empty() : new CellValue.Date(value);
        return this;
    }

    /**
     * Sets the cell value to a formula.
     */
    public Cell setFormula(@Nullable String expression) {
        if (expression == null) {
            this.value = new CellValue.Empty();
        } else {
            this.value = new CellValue.Formula(expression);
        }
        return this;
    }

    /**
     * Sets the cell to empty.
     */
    public Cell setEmpty() {
        this.value = new CellValue.Empty();
        return this;
    }

    /**
     * Sets the cell value directly.
     */
    public Cell setValue(@Nullable CellValue value) {
        this.value = value == null ? new CellValue.Empty() : value;
        return this;
    }

    /**
     * Sets the style ID.
     */
    public Cell style(int styleId) {
        this.styleId = styleId;
        return this;
    }

    /**
     * Returns the style ID.
     */
    public int styleId() {
        return styleId;
    }

    /**
     * Returns the cell value as a raw object (String, Double, Boolean, LocalDateTime, or null).
     */
    public @Nullable Object rawValue() {
        return switch (value) {
            case CellValue.Empty _ -> null;
            case CellValue.Text t -> t.value();
            case CellValue.Number n -> n.value();
            case CellValue.Bool b -> b.value();
            case CellValue.Date d -> d.value();
            case CellValue.Formula f -> f.expression();
            case CellValue.Error e -> e.code();
        };
    }

    @Override
    public String toString() {
        return String.format("Cell[col=%d, type=%s, value=%s]", column, value.type(), rawValue());
    }
}
