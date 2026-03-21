package com.beingidly.litexl;

/**
 * Represents the type of value stored in a cell.
 */
public enum CellType {
    /** Empty cell with no value. */
    EMPTY,
    /** Cell containing a string value. */
    STRING,
    /** Cell containing a numeric value. */
    NUMBER,
    /** Cell containing a boolean value. */
    BOOLEAN,
    /** Cell containing a date value. */
    DATE,
    /** Cell containing a formula. */
    FORMULA,
    /** Cell containing an error. */
    ERROR
}
