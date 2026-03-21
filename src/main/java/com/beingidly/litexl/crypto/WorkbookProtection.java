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
     */
    public static WorkbookProtection defaults() {
        return new WorkbookProtection(true, false);
    }

    /**
     * Returns a builder for customizing protection.
     */
    public static Builder builder() {
        return new Builder();
    }

    public static class Builder {
        private boolean lockStructure = true;
        private boolean lockWindows = false;

        public Builder lockStructure(boolean value) {
            this.lockStructure = value;
            return this;
        }

        public Builder lockWindows(boolean value) {
            this.lockWindows = value;
            return this;
        }

        public WorkbookProtection build() {
            return new WorkbookProtection(lockStructure, lockWindows);
        }
    }
}
