package com.beingidly.litexl.crypto;

/**
 * Options for sheet protection.
 */
public record SheetProtection(
    boolean selectLockedCells,
    boolean selectUnlockedCells,
    boolean formatCells,
    boolean formatColumns,
    boolean formatRows,
    boolean insertRows,
    boolean insertColumns,
    boolean deleteRows,
    boolean deleteColumns,
    boolean sort,
    boolean autoFilter,
    boolean pivotTables
) {
    /**
     * Returns default protection settings.
     */
    public static SheetProtection defaults() {
        return new SheetProtection(
            true,   // selectLockedCells
            true,   // selectUnlockedCells
            false,  // formatCells
            false,  // formatColumns
            false,  // formatRows
            false,  // insertRows
            false,  // insertColumns
            false,  // deleteRows
            false,  // deleteColumns
            false,  // sort
            false,  // autoFilter
            false   // pivotTables
        );
    }

    /**
     * Returns a builder for customizing protection.
     */
    public static Builder builder() {
        return new Builder();
    }

    public static class Builder {
        private boolean selectLockedCells = true;
        private boolean selectUnlockedCells = true;
        private boolean formatCells = false;
        private boolean formatColumns = false;
        private boolean formatRows = false;
        private boolean insertRows = false;
        private boolean insertColumns = false;
        private boolean deleteRows = false;
        private boolean deleteColumns = false;
        private boolean sort = false;
        private boolean autoFilter = false;
        private boolean pivotTables = false;

        public Builder selectLockedCells(boolean value) {
            this.selectLockedCells = value;
            return this;
        }

        public Builder selectUnlockedCells(boolean value) {
            this.selectUnlockedCells = value;
            return this;
        }

        public Builder formatCells(boolean value) {
            this.formatCells = value;
            return this;
        }

        public Builder formatColumns(boolean value) {
            this.formatColumns = value;
            return this;
        }

        public Builder formatRows(boolean value) {
            this.formatRows = value;
            return this;
        }

        public Builder insertRows(boolean value) {
            this.insertRows = value;
            return this;
        }

        public Builder insertColumns(boolean value) {
            this.insertColumns = value;
            return this;
        }

        public Builder deleteRows(boolean value) {
            this.deleteRows = value;
            return this;
        }

        public Builder deleteColumns(boolean value) {
            this.deleteColumns = value;
            return this;
        }

        public Builder sort(boolean value) {
            this.sort = value;
            return this;
        }

        public Builder autoFilter(boolean value) {
            this.autoFilter = value;
            return this;
        }

        public Builder pivotTables(boolean value) {
            this.pivotTables = value;
            return this;
        }

        public SheetProtection build() {
            return new SheetProtection(
                selectLockedCells,
                selectUnlockedCells,
                formatCells,
                formatColumns,
                formatRows,
                insertRows,
                insertColumns,
                deleteRows,
                deleteColumns,
                sort,
                autoFilter,
                pivotTables
            );
        }
    }
}
