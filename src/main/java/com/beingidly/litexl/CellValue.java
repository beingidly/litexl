package com.beingidly.litexl;

import java.time.LocalDateTime;

/**
 * Represents a cell value with type safety using sealed interface.
 */
public sealed interface CellValue {

    record Empty() implements CellValue {
        @Override
        public CellType type() {
            return CellType.EMPTY;
        }
    }

    record Text(String value) implements CellValue {
        @Override
        public CellType type() {
            return CellType.STRING;
        }
    }

    record Number(double value) implements CellValue {
        @Override
        public CellType type() {
            return CellType.NUMBER;
        }
    }

    record Bool(boolean value) implements CellValue {
        @Override
        public CellType type() {
            return CellType.BOOLEAN;
        }
    }

    record Date(LocalDateTime value) implements CellValue {
        @Override
        public CellType type() {
            return CellType.DATE;
        }
    }

    record Formula(String expression, CellValue cached) implements CellValue {
        public Formula(String expression) {
            this(expression, new Empty());
        }

        @Override
        public CellType type() {
            return CellType.FORMULA;
        }
    }

    record Error(String code) implements CellValue {
        @Override
        public CellType type() {
            return CellType.ERROR;
        }
    }

    /**
     * Returns the type of this cell value.
     */
    CellType type();
}
