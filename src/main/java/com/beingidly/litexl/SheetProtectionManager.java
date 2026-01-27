package com.beingidly.litexl;

import com.beingidly.litexl.crypto.SheetHasher;
import com.beingidly.litexl.crypto.SheetHasher.SheetProtectionInfo;
import com.beingidly.litexl.crypto.SheetProtection;
import org.jspecify.annotations.Nullable;

import java.util.Arrays;

/**
 * Manages protection settings for a worksheet.
 *
 * <p>This class handles sheet protection with password hashing.
 * Passwords are accepted as char[] and immediately hashed for security -
 * the password itself is never stored.
 *
 * <p>This class is <b>not thread-safe</b>.
 * External synchronization is required for concurrent access.
 */
public final class SheetProtectionManager {

    private static final SheetHasher HASHER = new SheetHasher();

    private SheetProtection options;
    private SheetProtectionInfo passwordInfo;

    SheetProtectionManager() {
    }

    /**
     * Protects the sheet with a password and options.
     *
     * <p>The password is immediately hashed and the char array is cleared.
     * The password itself is never stored.
     *
     * @param password the password (will be cleared after hashing), or null for no password
     * @param options the protection options
     */
    public void protect(char @Nullable [] password, SheetProtection options) {
        this.options = options;
        if (password != null && password.length > 0) {
            try {
                this.passwordInfo = HASHER.hash(new String(password));
            } finally {
                Arrays.fill(password, '\0');
            }
        } else {
            this.passwordInfo = null;
        }
    }

    /**
     * Protects the sheet without a password.
     *
     * @param options the protection options
     */
    public void protect(SheetProtection options) {
        this.options = options;
        this.passwordInfo = null;
    }

    /**
     * Unprotects the sheet if the password matches.
     *
     * <p>The password char array is cleared after verification.
     *
     * @param password the password to verify (will be cleared after verification)
     * @return true if unprotected successfully
     */
    public boolean unprotect(char @Nullable [] password) {
        if (this.options == null) {
            return true;
        }

        if (this.passwordInfo == null) {
            this.options = null;
            return true;
        }

        try {
            if (HASHER.verify(new String(password), this.passwordInfo)) {
                this.options = null;
                this.passwordInfo = null;
                return true;
            }
            return false;
        } finally {
            if (password != null) {
                Arrays.fill(password, '\0');
            }
        }
    }

    /**
     * Unprotects the sheet (no password required if sheet has no password).
     *
     * @return true if unprotected successfully
     */
    public boolean unprotect() {
        if (this.passwordInfo == null) {
            this.options = null;
            return true;
        }
        return false;
    }

    /**
     * Returns the protection options, or null if not protected.
     */
    public @Nullable SheetProtection options() {
        return options;
    }

    /**
     * Returns true if the sheet is protected.
     */
    public boolean isProtected() {
        return options != null;
    }

    /**
     * Returns the password info for internal use during save.
     * Returns null if no password was set.
     */
    @Nullable SheetProtectionInfo passwordInfo() {
        return passwordInfo;
    }

    /**
     * Sets the password info directly (for internal use when reading).
     */
    void setPasswordInfo(@Nullable SheetProtectionInfo info) {
        this.passwordInfo = info;
    }
}
