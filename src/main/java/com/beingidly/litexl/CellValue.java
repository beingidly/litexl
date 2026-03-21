package com.beingidly.litexl;

import java.time.LocalDateTime;

/**
 * Represents a cell value with type safety using sealed interface.
 */
public sealed interface CellValue {

    /** An empty cell value. */
    record Empty() implements CellValue {
        @Override
        public CellType type() {
            return CellType.EMPTY;
        }
    }

    /**
     * A string cell value.
     *
     * @param value the string content
     */
    record Text(String value) implements CellValue {
        @Override
        public CellType type() {
            return CellType.STRING;
        }
    }

    /**
     * A numeric cell value.
     *
     * @param value the numeric content
     */
    record Number(double value) implements CellValue {
        @Override
        public CellType type() {
            return CellType.NUMBER;
        }
    }

    /**
     * A boolean cell value.
     *
     * @param value the boolean content
     */
    record Bool(boolean value) implements CellValue {
        @Override
        public CellType type() {
            return CellType.BOOLEAN;
        }
    }

    /**
     * A date cell value.
     *
     * @param value the date-time content
     */
    record Date(LocalDateTime value) implements CellValue {
        @Override
        public CellType type() {
            return CellType.DATE;
        }
    }

    /**
     * A formula cell value.
     *
     * @param expression the formula expression
     * @param cached the cached result value
     */
    record Formula(String expression, CellValue cached) implements CellValue {
        /**
         * Creates a formula with no cached value.
         *
         * @param expression the formula expression
         */
        public Formula(String expression) {
            this(expression, new Empty());
        }

        @Override
        public CellType type() {
            return CellType.FORMULA;
        }
    }

    /**
     * An error cell value.
     *
     * @param code the error code
     */
    record Error(String code) implements CellValue {
        @Override
        public CellType type() {
            return CellType.ERROR;
        }
    }

    /**
     * Returns the type of this cell value.
     *
     * @return the cell type
     */
    CellType type();
}
