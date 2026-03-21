package com.beingidly.litexl.crypto;

/**
 * Options for sheet protection.
 *
 * @param selectLockedCells whether users can select locked cells
 * @param selectUnlockedCells whether users can select unlocked cells
 * @param formatCells whether users can format cells
 * @param formatColumns whether users can format columns
 * @param formatRows whether users can format rows
 * @param insertRows whether users can insert rows
 * @param insertColumns whether users can insert columns
 * @param deleteRows whether users can delete rows
 * @param deleteColumns whether users can delete columns
 * @param sort whether users can sort
 * @param autoFilter whether users can use auto-filter
 * @param pivotTables whether users can use pivot tables
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
     *
     * @return the default protection settings
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
     *
     * @return a new builder
     */
    public static Builder builder() {
        return new Builder();
    }

    /** Builder for customizing sheet protection settings. */
    public static class Builder {
        /** Creates a new builder with default settings. */
        public Builder() {}
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

        /**
         * Sets locked cell selection.
         * @param value the value
         * @return this builder
         */
        public Builder selectLockedCells(boolean value) {
            this.selectLockedCells = value;
            return this;
        }

        /**
         * Sets unlocked cell selection.
         * @param value the value
         * @return this builder
         */
        public Builder selectUnlockedCells(boolean value) {
            this.selectUnlockedCells = value;
            return this;
        }

        /**
         * Sets cell formatting permission.
         * @param value the value
         * @return this builder
         */
        public Builder formatCells(boolean value) {
            this.formatCells = value;
            return this;
        }

        /**
         * Sets column formatting permission.
         * @param value the value
         * @return this builder
         */
        public Builder formatColumns(boolean value) {
            this.formatColumns = value;
            return this;
        }

        /**
         * Sets row formatting permission.
         * @param value the value
         * @return this builder
         */
        public Builder formatRows(boolean value) {
            this.formatRows = value;
            return this;
        }

        /**
         * Sets row insertion permission.
         * @param value the value
         * @return this builder
         */
        public Builder insertRows(boolean value) {
            this.insertRows = value;
            return this;
        }

        /**
         * Sets column insertion permission.
         * @param value the value
         * @return this builder
         */
        public Builder insertColumns(boolean value) {
            this.insertColumns = value;
            return this;
        }

        /**
         * Sets row deletion permission.
         * @param value the value
         * @return this builder
         */
        public Builder deleteRows(boolean value) {
            this.deleteRows = value;
            return this;
        }

        /**
         * Sets column deletion permission.
         * @param value the value
         * @return this builder
         */
        public Builder deleteColumns(boolean value) {
            this.deleteColumns = value;
            return this;
        }

        /**
         * Sets sort permission.
         * @param value the value
         * @return this builder
         */
        public Builder sort(boolean value) {
            this.sort = value;
            return this;
        }

        /**
         * Sets auto-filter permission.
         * @param value the value
         * @return this builder
         */
        public Builder autoFilter(boolean value) {
            this.autoFilter = value;
            return this;
        }

        /**
         * Sets pivot table permission.
         * @param value the value
         * @return this builder
         */
        public Builder pivotTables(boolean value) {
            this.pivotTables = value;
            return this;
        }
        /**
         * Builds the protection settings.
         * @return a new SheetProtection
         */
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
