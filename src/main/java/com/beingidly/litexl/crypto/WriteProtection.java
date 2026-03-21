package com.beingidly.litexl.crypto;

/**
 * Write protection settings for a workbook (maps to the {@code fileSharing} element in workbook.xml).
 *
 * <p>When set, Excel recommends opening the file as read-only and optionally
 * requires a password to modify the file.</p>
 *
 * @param readOnlyRecommended whether to recommend opening as read-only
 * @param userName the name of the user who set the protection
 */
public record WriteProtection(
    boolean readOnlyRecommended,
    String userName
) {
    /** Validates that userName is not null. */
    public WriteProtection {
        if (userName == null) {
            throw new IllegalArgumentException("userName cannot be null");
        }
    }
}
