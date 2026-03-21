package com.beingidly.litexl.crypto;

/**
 * Options for workbook structure protection (maps to the {@code workbookProtection} element in workbook.xml).
 *
 * <p>Workbook protection prevents structural changes to the workbook such as
 * adding, deleting, or renaming sheets.</p>
 *
 * @param lockStructure whether to prevent adding, deleting, or renaming sheets
 * @param lockWindows whether to prevent changing window size and position
 */
public record WorkbookProtection(
    boolean lockStructure,
    boolean lockWindows
) {
    /**
     * Returns default protection settings (structure locked, windows unlocked).
     *
     * @return the default protection settings
     */
    public static WorkbookProtection defaults() {
        return new WorkbookProtection(true, false);
    }

    /**
     * Returns a builder for customizing protection.
     *
     * @return a new builder
     */
    public static Builder builder() {
        return new Builder();
    }

    /** Builder for customizing workbook protection settings. */
    public static class Builder {
        /** Creates a new builder with default settings. */
        public Builder() {}
        private boolean lockStructure = true;
        private boolean lockWindows = false;

        /**
         * Sets structure locking.
         * @param value true to lock structure
         * @return this builder
         */
        public Builder lockStructure(boolean value) {
            this.lockStructure = value;
            return this;
        }

        /**
         * Sets window locking.
         * @param value true to lock windows
         * @return this builder
         */
        public Builder lockWindows(boolean value) {
            this.lockWindows = value;
            return this;
        }

        /**
         * Builds the protection settings.
         * @return a new WorkbookProtection
         */
        public WorkbookProtection build() {
            return new WorkbookProtection(lockStructure, lockWindows);
        }
    }
}
