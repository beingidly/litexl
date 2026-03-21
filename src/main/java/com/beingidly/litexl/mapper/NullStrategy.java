package com.beingidly.litexl.mapper;

/**
 * Strategy for handling null values during mapping.
 */
public enum NullStrategy {
    /** Skip null values (do not write a cell). */
    SKIP,
    /** Write an empty cell for null values. */
    EMPTY_CELL
}
